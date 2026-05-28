import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io

st.set_page_config(page_title="搭售分析看板 & 宽表生成器", layout="wide")

st.title("📊 搭售分析明细宽表生成系统")
st.markdown("---")

# ==========================================
# 1. 侧边栏：文件上传区
# ==========================================
st.sidebar.header("📁 数据源上传")
file_b_raw = st.sidebar.file_uploader("1. 上传 数据底表 (CSV 或 Excel)", type=["csv", "xlsx"])
file_a_raw = st.sidebar.file_uploader("2. 上传 维度A整体映射表 (Excel)", type=["xlsx"])

# 在 Session State 中持久化存储生成的宽表，避免切换图表时网页刷新重新计算
if "wide_final" not in st.session_state:
    st.session_state.wide_final = None

# ==========================================
# 2. 核心数据处理引擎
# ==========================================
if file_b_raw and file_a_raw:
    if st.sidebar.button("🚀 开始生成宽表并校准逻辑"):
        with st.spinner("数据清洗中，请稍候..."):
            try:
                # 读取维度A
                df_a = pd.read_excel(file_a_raw)
                
                # 读取维度B (底表)，动态兼容CSV和Excel
                if file_b_raw.name.endswith('.csv'):
                    df_b = pd.read_csv(file_b_raw)
                else:
                    df_b = pd.read_excel(file_b_raw)
                
                # 清理首尾空格
                df_a.columns = [str(c).strip() for c in df_a.columns]
                df_b.columns = [str(c).strip() for c in df_b.columns]
                
                # ---- 2.1 预处理维度 A ----
                a_cols_map = {
                    '销售月份(month)': '销售月份',
                    '销售国': '站点',
                    'yspu_code': '搭售件编码',
                    '总销量': '搭售件总销量',
                    '单独销量': '搭售件单独销量',
                    '搭售销量': '搭售件搭售销量',
                    '适用主销品总销量': '适用主销品总销量',
                    '搭售金额': '搭售销售金额',
                    '单独销售金额': '单独销售金额'
                }
                df_a_unique = df_a[list(a_cols_map.keys())].rename(columns=a_cols_map)
                
                # ---- 2.2 预处理维度 B & 剥离主销品大盘 ----
                df_main_only = df_b[df_b['产品类型'] == '主销品'].copy()
                df_main_product_sales = df_main_only[['销售月份(month)', '销售国', 'yspu_code', '销量']].copy()
                df_main_product_sales.rename(
                    columns={'yspu_code': '被搭售的主销品yspu_code', '销量': '主销品单品销量'}, 
                    inplace=True
                )
                df_main_product_sales = df_main_product_sales.groupby(
                    ['销售月份(month)', '销售国', '被搭售的主销品yspu_code']
                )['主销品单品销量'].sum().reset_index()
                
                # 过滤掉配件单独销售（剔除被搭售主销品为'-'或空值的行）
                main_code_col = 'get被搭售的主销品yspu_code' if 'get被搭售的主销品yspu_code' in df_b.columns else '被搭售的主销品yspu_code'
                df_b_items = df_b[
                    (df_b['产品类型'] == '搭售件') & 
                    (df_b[main_code_col] != '-') & 
                    (df_b[main_code_col].notna())
                ].copy()
                
                # 合并主销品大盘销量
                df_b_with_true_sales = pd.merge(
                    df_b_items, df_main_product_sales,
                    left_on=['销售月份(month)', '销售国', main_code_col],
                    right_on=['销售月份(month)', '销售国', '被搭售的主销品yspu_code'],
                    how='left'
                )
                
                b_cols_map = {
                    '销售月份(month)': '销售月份',
                    '销售国': '站点',
                    'yspu': '搭售件名称',
                    'yspu_code': '搭售件编码',
                    '搭售销量': '主销品带动搭售量',
                    '被搭售的主销品yspu': '被搭售的主销品yspu',
                    '被搭售的主销品yspu_code': '被搭售的主销品yspu_code',
                    '主销品单品销量': '主销品单品销量',   
                    '场景': '场景'
                }
                df_b_core = df_b_with_true_sales[list(b_cols_map.keys())].rename(columns=b_cols_map)
                
                # ---- 2.3 最终全局合并 ----
                wide = pd.merge(df_b_core, df_a_unique, on=['销售月份', '站点', '搭售件编码'], how='left')
                
                # ---- 2.4 【新需求】派生指标计算 (安全除法，防0报错) ----
                wide['占比'] = wide['搭售件搭售销量'] / wide['搭售件总销量']
                wide['整体搭售率'] = wide['搭售件搭售销量'] / wide['适用主销品总销量']
                wide['明细搭售率'] = wide['主销品带动搭售量'] / wide['主销品单品销量']
                
                # 统一清理下空值或除零产生的 inf / nan
                for rate_col in ['占比', '整体搭售率', '明细搭售率']:
                    wide[rate_col] = wide[rate_col].fillna(0).replace([float('inf'), float('-inf')], 0)
                
                # ---- 2.5 严格控制输出列顺序 ----
                final_order = [
                    '销售月份', '站点', '搭售件名称', '搭售件编码', 
                    '搭售件总销量', '搭售件单独销量', '搭售件搭售销量',
                    '单独销售金额', '搭售销售金额', '适用主销品总销量',
                    '被搭售的主销品yspu', '被搭售的主销品yspu_code', 
                    '主销品单品销量', '主销品带动搭售量', '场景',
                    '占比', '整体搭售率', '明细搭售率'
                ]
                existing_cols = [c for c in final_order if c in wide.columns]
                st.session_state.wide_final = wide[existing_cols]
                st.success("🎉 宽表计算成功！已自动校准大盘销量并生成衍生指标。")
                
            except Exception as e:
                st.error(f"❌ 数据处理冲突，请核对字段格式。报错信息: {e}")

# ==========================================
# 3. 结果下载与可视化呈现
# ==========================================
if st.session_state.wide_final is not None:
    df_res = st.session_state.wide_final
    
    # ---- 3.1 提供下载按钮 ----
    st.subheader("📥 下载中心")
    # 将 DataFrame 转换为内存中的 Excel 对象
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_res.to_excel(writer, index=False, sheet_name='搭售明细宽表')
    
    st.download_button(
        label="💾 点击下载最终『搭售分析明细宽表.xlsx』",
        data=buffer.getvalue(),
        file_name="搭售分析明细宽表_最终版.xlsx",
        mime="application/vnd.ms-excel"
    )
    
    # 预览数据
    with st.expander("🔍 展开查看宽表实时数据预览 (前 50 条)"):
        st.dataframe(df_res.head(50))
    
    st.markdown("---")
    
    # ---- 3.2 动态可视化看板 ----
    st.subheader("📈 搭售率趋势动态可视化")
    
    # 业务多选/单选框：支持模糊搜索和选择搭售件
    yspu_list = sorted(df_res['搭售件名称'].unique().tolist())
    selected_yspu = st.selectbox("🎯 请选择或输入要分析的搭售件 (YSPU)：", yspu_list)
    
    if selected_yspu:
        # 联动过滤出当前搭售件的数据
        df_chart = df_res[df_res['搭售件名称'] == selected_yspu].copy()
        
        # 按照时间、国家排序确保折线连续
        df_chart['排序键'] = df_chart['销售月份'].astype(str) + "_" + df_chart['站点'].astype(str)
        df_chart = df_chart.sort_values(by='排序键')
        
        # 为了保证鼠标悬浮能清晰看到是在哪个国家的哪个主销品带动的，拼接 hover 信息
        hover_texts = []
        for idx, row in df_chart.iterrows():
            text = (
                f"<b>月份-站点:</b> {row['销售月份']}-{row['站点']}<br>"
                f"<b>关联主销品:</b> {row['被搭售的主销品yspu']}<br>"
                f"<b>搭售件单独销量:</b> {row['搭售件单独销量']}<br>"
                f"<b>主销品单品销量:</b> {row['主销品单品销量']}<br>"
                f"<b>主销品带动搭售量:</b> {row['主销品带动搭售量']}"
            )
            hover_texts.append(text)
            
        # 使用 Plotly 绘制高交互线图（原生支持悬停和缩放）
        fig = go.Figure()
        
        # 整体搭售率折线
        fig.add_trace(go.Scatter(
            x=df_chart['排序键'],
            y=df_chart['整体搭售率'],
            mode='lines+markers',
            name='整体搭售率',
            line=dict(color='#1f77b4', width=3),
            text=hover_texts,
            hoverinfo='text+y'
        ))
        
        # 明细搭售率折线
        fig.add_trace(go.Scatter(
            x=df_chart['排序键'],
            y=df_chart['明细搭售率'],
            mode='lines+markers',
            name='明细搭售率',
            line=dict(color='#ff7f0e', width=2, dash='dash'),
            text=hover_texts,
            hoverinfo='text+y'
        ))
        
        fig.update_layout(
            title=dict(text=f"📊 {selected_yspu} 的搭售率健康度走势图", font=dict(size=18)),
            xaxis_title="业务时间轴 (月份_国家)",
            yaxis_title="比率 (Percentage)",
            yaxis=dict(tickformat=".2%"), # 自动转化为百分比显示
            hovermode="closest",
            legend=dict(orient="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            template="streamlit"
        )
        
        # 渲染图表
        st.plotly_chart(fig, use_container_width=True)

else:
    st.info("💡 提示：请在左侧栏依次上传『数据底表』和『整体映射表』，然后点击开始按钮。生成的完美宽表和趋势图将在此处动态呈现。")
