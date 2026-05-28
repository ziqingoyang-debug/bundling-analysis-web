import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io

st.set_page_config(page_title="搭售分析看板 & 宽表生成器", layout="wide")

st.title("📊 搭售分析明细宽表生成系统")
st.markdown("---")

# ==========================================
# 1. 侧边栏：文件上传区 (完全对齐您的真实附件)
# ==========================================
st.sidebar.header("📁 数据源上传")
file_b_raw = st.sidebar.file_uploader("1. 上传 IT提供搭售数据底表 (CSV 或 Excel)", type=["csv", "xlsx"])
file_map_raw = st.sidebar.file_uploader("2. 上传 产品-场景套系映射表 (Excel 或 CSV)", type=["xlsx", "csv"])

if "wide_final" not in st.session_state:
    st.session_state.wide_final = None

# ==========================================
# 2. 核心数据处理引擎 (底表清洗 -> 映射表匹配 -> 衍生指标纯代码计算)
# ==========================================
if file_b_raw and file_map_raw:
    if st.sidebar.button("🚀 开始清洗并生成搭售分析明细宽表"):
        with st.spinner("数分引擎正在计算整体与明细搭售率，请稍候..."):
            try:
                # 2.1 动态读取数据底表
                if file_b_raw.name.endswith('.csv'):
                    df_b = pd.read_csv(file_b_raw)
                else:
                    df_b = pd.read_excel(file_b_raw)
                
                # 2.2 动态读取产品场景映射表
                if file_map_raw.name.endswith('.csv'):
                    df_map = pd.read_csv(file_map_raw)
                else:
                    df_map = pd.read_excel(file_map_raw)
                
                # 清理表头空格
                df_b.columns = [str(c).strip() for c in df_b.columns]
                df_map.columns = [str(c).strip() for c in df_map.columns]
                
                # ----------------------------------------------------
                # 💡 第一步：从底表中提取【主销品大盘销量】库
                # ----------------------------------------------------
                df_main_pool = df_b[df_b['产品类型'] == '主销品'].copy()
                df_main_sum = df_main_pool.groupby(['销售月份(month)', '销售国', 'yspu_code'])['销量'].sum().reset_index()
                df_main_sum.rename(columns={'yspu_code': '主销品编码', '销量': '大盘主销品单品销量'}, inplace=True)
                
                # ----------------------------------------------------
                # 💡 第二步：纯代码计算【搭售件的大盘指标】（总销量、单独销量、搭售销量、单独/搭售金额）
                # ----------------------------------------------------
                df_acc_pool = df_b[df_b['产品类型'] == '搭售件'].copy()
                
                # 计算金额列（清洗数据里的'万'字并转为数值）
                for amt_col in ['单独销售金额_cny', '搭售金额_cny']:
                    if amt_col in df_acc_pool.columns:
                        df_acc_pool[amt_col] = df_acc_pool[amt_col].astype(str).str.replace('万', '').astype(float) * 10000 if df_acc_pool[amt_col].dtype == object else df_acc_pool[amt_col]
                    else:
                        df_acc_pool[amt_col] = 0
                
                # 聚合出搭售件在当月当国的全局大盘指标
                df_acc_market = df_acc_pool.groupby(['销售月份(month)', '销售国', 'yspu_code']).agg({
                    '销量': 'sum',
                    '单独销量': 'sum',
                    '搭售销量': 'sum',
                    '单独销售金额_cny': 'sum',
                    '搭售金额_cny': 'sum'
                }).reset_index().rename(columns={
                    '销量': '搭售件总销量',
                    '单独销量': '搭售件单独销量',
                    '搭售销量': '搭售件搭售销量',
                    '单独销售金额_cny': '单独销售金额',
                    '搭售金额_cny': '搭售销售金额'
                })

                # ----------------------------------------------------
                # 💡 第三步：解析《产品-场景套系映射表》，计算【适用主销品总销量】
                # ----------------------------------------------------
                # 建立一个空字典，用来存“配件编码 -> 适用主销品编码列表”的映射
                map_dict = {}
                for _, row in df_map.iterrows():
                    acc_code = str(row['搭售件YSPU']).strip()
                    main_codes_str = str(row['能被搭售品YSPU'] if '能被搭售品YSPU' in df_map.columns else row['能被搭售的主销品YSPU']).strip()
                    # 拆分成列表
                    main_codes_list = [c.strip() for c in main_codes_str.split(',') if c.strip()]
                    map_dict[acc_code] = main_codes_list
                
                # 计算每个配件在特定月份和国家的【适用主销品总销量】
                applicable_sales_list = []
                for _, row in df_acc_market.iterrows():
                    m_month = row['销售月份(month)']
                    m_site = row['销售国']
                    m_acc = row['yspu_code']
                    
                    app_main_codes = map_dict.get(m_acc, [])
                    # 去主销品大盘库里筛选出这些主销品
                    sub_main = df_main_sum[
                        (df_main_sum['销售月份(month)'] == m_month) & 
                        (df_main_sum['销售国'] == m_site) & 
                        (df_main_sum['主销品编码'].isin(app_main_codes))
                    ]
                    total_app_main_sales = sub_main['大盘主销品单品销量'].sum()
                    applicable_sales_list.append(total_app_main_sales)
                
                df_acc_market['适用主销品总销量'] = applicable_sales_list

                # ----------------------------------------------------
                # 💡 第四步：清洗底表，拆解出【真实的明细行】（剔除"-"）并合并所有指标
                # ----------------------------------------------------
                # 动态识别IT底表中的被搭售主销品编码列
                main_code_col = 'get被搭售的主销品yspu_code' if 'get被搭售的主销品yspu_code' in df_b.columns else '被搭售的主销品yspu_code'
                
                # 过滤掉配件单独销售（剔除被搭售主销品为'-'或空值的行，锁定真正明细）
                df_b_detail = df_b[
                    (df_b['产品类型'] == '搭售件') & 
                    (df_b[main_code_col] != '-') & 
                    (df_b[main_code_col].notna())
                ].copy()
                
                # 1. 挂载主销品单品销量
                df_b_detail = pd.merge(
                    df_b_detail, df_main_sum,
                    left_on=['销售月份(month)', '销售国', main_code_col],
                    right_on=['销售月份(month)', '销售国', '主销品编码'],
                    how='left'
                )
                
                # 2. 挂载搭售件的大盘指标和适用主销品总销量
                wide = pd.merge(
                    df_b_detail, df_acc_market,
                    on=['销售月份(month)', '销售国', 'yspu_code'],
                    how='left'
                )
                
                # ----------------------------------------------------
                # 💡 第五步：标准重命名与三个衍生率指标计算
                # ----------------------------------------------------
                wide['占比'] = wide['搭售件搭售销量'] / wide['搭售件总销量']
                wide['整体搭售率'] = wide['搭售件搭售销量'] / wide['适用主销品总销量']
                wide['明细搭售率'] = wide['搭售销量_x'] / wide['大盘主销品单品销量']
                
                # 清理空值与除零错
                for rate_col in ['占比', '整体搭售率', '明细搭售率']:
                    wide[rate_col] = wide[rate_col].fillna(0).replace([float('inf'), float('-inf')], 0)
                
                # 构建最终的标准宽表
                df_final = pd.DataFrame({
                    '销售月份': wide['销售月份(month)'],
                    '站点': wide['销售国'],
                    '搭售件名称': wide['yspu'],
                    '搭售件编码': wide['yspu_code'],
                    '搭售件总销量': wide['搭售件总销量'],
                    '搭售件单独销量': wide['搭售件单独销量'],
                    '搭售件搭售销量': wide['搭售件搭售销量'],
                    '单独销售金额': wide['单独销售金额'],
                    '搭售销售金额': wide['搭售销售金额'],
                    '适用主销品总销量': wide['适用主销品总销量'],
                    '被搭售的主销品yspu': wide['被搭售的主销品yspu'],
                    '被搭售的主销品yspu_code': wide[main_code_col],
                    '主销品单品销量': wide['大盘主销品单品销量'],
                    '主销品带动搭售量': wide['搭售销量_x'],
                    '场景': wide['场景'] if '场景' in wide.columns else '基础',
                    '占比': wide['占比'],
                    '整体搭售率': wide['整体搭售率'],
                    '明细搭售率': wide['明细搭售率']
                })
                
                st.session_state.wide_final = df_final
                st.success("🎉 宽表及衍生指标计算成功！大盘数据已通过代码底层完美自动对齐。")
                
            except Exception as e:
                st.error(f"❌ 数据清洗冲突，请核对字段格式。报错信息: {e}")

# ==========================================
# 3. 结果下载与动态线图可视化 Presentation
# ==========================================
if st.session_state.wide_final is not None:
    df_res = st.session_state.wide_final
    
    st.subheader("📥 下载中心")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_res.to_excel(writer, index=False, sheet_name='搭售明细宽表')
    
    st.download_button(
        label="💾 点击下载最终『搭售分析明细宽表.xlsx』",
        data=buffer.getvalue(),
        file_name="搭售分析明细宽表_最终版.xlsx",
        mime="application/vnd.ms-excel"
    )
    
    with st.expander("🔍 展开查看宽表实时数据预览"):
        st.dataframe(df_res.head(30))
    
    st.markdown("---")
    st.subheader("📈 搭售率趋势动态可视化")
    
    # 支持模糊搜索和选择搭售件名称
    yspu_list = sorted(df_res['搭售件名称'].unique().tolist())
    selected_yspu = st.selectbox("🎯 请选择或输入要分析的搭售件 (YSPU)：", yspu_list)
    
    if selected_yspu:
        df_chart = df_res[df_res['搭售件名称'] == selected_yspu].copy()
        
        # 按照时间、国家排序确保折线连续
        df_chart['排序键'] = df_chart['销售月份'].astype(str) + "_" + df_chart['站点'].astype(str)
        df_chart = df_chart.sort_values(by='排序键')
        
        # 拼接高交互鼠标悬停 hover 信息
        hover_texts = []
        for idx, row in df_chart.iterrows():
            text = (
                f"<b>月份-站点:</b> {row['销售月份']}-{row['站点']}<br>"
                f"<b>关联主销品:</b> {row['被搭售的主销品yspu']}<br>"
                f"<b>主销品单品销量:</b> {int(row['主销品单品销量'])}<br>"
                f"<b>主销品带动搭售量:</b> {int(row['主销品带动搭售量'])}"
            )
            hover_texts.append(text)
            
        fig = go.Figure()
        
        # 1. 整体搭售率折线 (粗实线)
        fig.add_trace(go.Scatter(
            x=df_chart['排序键'], y=df_chart['整体搭售率'],
            mode='lines+markers', name='整体搭售率',
            line=dict(color='#1f77b4', width=3), text=hover_texts, hoverinfo='text+y'
        ))
        
        # 2. 明细搭售率折线 (虚线)
        fig.add_trace(go.Scatter(
            x=df_chart['排序键'], y=df_chart['明细搭售率'],
            mode='lines+markers', name='明细搭售率',
            line=dict(color='#ff7f0e', width=2, dash='dash'), text=hover_texts, hoverinfo='text+y'
        ))
        
        fig.update_layout(
            title=dict(text=f"📊 {selected_yspu} 的搭售率走势分析（整体 vs 明细）", font=dict(size=18)),
            xaxis_title="时间轴与站点 (月份_国家)", yaxis_title="比率 (Percentage)",
            yaxis=dict(tickformat=".2%"), hovermode="closest",
            legend=dict(orient="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            template="streamlit"
        )
        st.plotly_chart(fig, use_container_width=True)
else:
    st.info("💡 提示：请在左侧栏上传【1. IT搭售数据底表】和【2. 产品-场景套系映射表】，然后点击按钮启动清洗。")
