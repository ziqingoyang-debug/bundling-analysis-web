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
file_b_raw = st.sidebar.file_uploader("1. 上传 IT提供搭售数据底表 (CSV 或 Excel)", type=["csv", "xlsx"])
file_map_raw = st.sidebar.file_uploader("2. 上传 产品-场景套系映射表 (Excel 或 CSV)", type=["xlsx", "csv"])

if "wide_final" not in st.session_state:
    st.session_state.wide_final = None

# ==========================================
# 2. 核心数据处理引擎
# ==========================================
if file_b_raw and file_map_raw:
    if st.sidebar.button("🚀 开始清洗并生成搭售分析明细宽表"):
        with st.spinner("数分引擎正在清洗整合大盘数据，请稍候..."):
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
                
                # 安全清理表头空格
                df_b.columns = [str(c).strip() for c in df_b.columns]
                df_map.columns = [str(c).strip() for c in df_map.columns]
                
                # ----------------------------------------------------
                # 💡 第一步：从底表中提取【大盘主销品单品销量】
                # ----------------------------------------------------
                df_main_pool = df_b[df_b['产品类型'] == '主销品'].copy()
                if len(df_main_pool) == 0:
                    df_main_pool = df_b[df_b['产品类型'].str.contains('主', na=False)].copy()
                    
                df_main_sum = df_main_pool.groupby(['销售月份(month)', '销售国', 'yspu_code'])['销量'].sum().reset_index()
                df_main_sum.rename(columns={'yspu_code': '匹配用_主销品编码', '销量': '大盘主销品单品销量'}, inplace=True)
                
                # ----------------------------------------------------
                # 💡 第二步：计算【搭售件的大盘指标】
                # ----------------------------------------------------
                df_acc_pool = df_b[df_b['产品类型'] == '搭售件'].copy()
                if len(df_acc_pool) == 0:
                    df_acc_pool = df_b[df_b['产品类型'].str.contains('搭|配', na=False)].copy()
                
                # 清洗金额列中的 '万' 字并转换为数值
                for amt_col in ['单独销售金额_cny', '搭售金额_cny']:
                    if amt_col in df_acc_pool.columns:
                        if df_acc_pool[amt_col].dtype == object:
                            df_acc_pool[amt_col] = df_acc_pool[amt_col].astype(str).str.replace('万', '').astype(float) * 10000
                        else:
                            df_acc_pool[amt_col] = df_acc_pool[amt_col].astype(float)
                    else:
                        df_acc_pool[amt_col] = 0.0
                
                # 聚合出大盘数据
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
                # 💡 第三步：安全解析映射表 (彻底修复 AttributeError)
                # ----------------------------------------------------
                map_dict = {}
                main_col_in_map = '能被搭售的主销品YSPU' if '能被搭售的主销品YSPU' in df_map.columns else '能被搭售品YSPU'
                
                for _, row in df_map.iterrows():
                    acc_code = str(row['搭售件YSPU']).strip()
                    # 强转为最基础的字符串，防止长文本对象错乱
                    main_codes_str = str(row[main_col_in_map])
                    # 替换可能带来Bug的双引号、括号或换行符
                    main_codes_str = main_codes_str.replace('"', '').replace('\n', '').replace('\r', '')
                    # 切割成干净的编码列表
                    main_codes_list = [c.strip() for c in main_codes_str.split(',') if c.strip()]
                    map_dict[acc_code] = main_codes_list
                
                applicable_sales_list = []
                for _, row in df_acc_market.iterrows():
                    m_month = row['销售月份(month)']
                    m_site = row['销售国']
                    m_acc = row['yspu_code']
                    
                    app_main_codes = map_dict.get(m_acc, [])
                    
                    # 💡 安全防护：过滤主销品销量，防止 DataFrame 属性错乱
                    if isinstance(df_main_sum, pd.DataFrame) and len(df_main_sum) > 0:
                        sub_main = df_main_sum[
                            (df_main_sum['销售月份(month)'] == m_month) & 
                            (df_main_sum['销售国'] == m_site) & 
                            (df_main_sum['匹配用_主销品编码'].isin(app_main_codes))
                        ]
                        val_sum = sub_main['大盘主销品单品销量'].sum() if len(sub_main) > 0 else 0
                    else:
                        val_sum = 0
                    applicable_sales_list.append(val_sum)
                
                df_acc_market['适用主销品总销量'] = applicable_sales_list

                # ----------------------------------------------------
                # 💡 第四步：拆解出【真实的明细行】
                # ----------------------------------------------------
                main_code_col = 'get被搭售的主销品yspu_code' if 'get被搭售的主销品yspu_code' in df_b.columns else '被搭售的主销品yspu_code'
                
                df_b_detail = df_b[
                    (df_b['产品类型'] == '搭售件') & 
                    (df_b[main_code_col] != '-') & 
                    (df_b[main_code_col].notna())
                ].copy()
                
                # 重新命名防冲突
                df_b_detail.rename(columns={'搭售销量': '主销品带动搭售量'}, inplace=True)
                
                # 1. 挂载主销品单品销量
                df_b_detail = pd.merge(
                    df_b_detail, df_main_sum,
                    left_on=['销售月份(month)', '销售国', main_code_col],
                    right_on=['销售月份(month)', '销售国', '匹配用_主销品编码'],
                    how='left'
                )
                df_b_detail['大盘主销品单品销量'] = df_b_detail['大盘主销品单品销量'].fillna(0)
                
                # 2. 挂载搭售件大盘及适用主销品总销量
                wide = pd.merge(
                    df_b_detail, df_acc_market,
                    on=['销售月份(month)', '销售国', 'yspu_code'],
                    how='left'
                )
                
                # ----------------------------------------------------
                # 💡 第五步：指标计算与最终宽表重构
                # ----------------------------------------------------
                wide['占比'] = wide['搭售件搭售销量'] / wide['搭售件总销量']
                wide['整体搭售率'] = wide['搭售件搭售销量'] / wide['适用主销品总销量']
                wide['明细搭售率'] = wide['主销品带动搭售量'] / wide['大盘主销品单品销量']
                
                for rate_col in ['占比', '整体搭售率', '明细搭售率']:
                    wide[rate_col] = wide[rate_col].fillna(0).replace([float('inf'), float('-inf')], 0)
                
                df_final = pd.DataFrame({
                    '销售月份': wide['销售月份(month)'],
                    '站点': wide['销售国'],
                    '搭售件名称': wide['yspu'],
                    '搭售件编码': wide['yspu_code'],
                    '搭售件总销量': wide['搭售件总销量'].fillna(0),
                    '搭售件单独销量': wide['搭售件单独销量'].fillna(0),
                    '搭售件搭售销量': wide['搭售件搭售销量'].fillna(0),
                    '单独销售金额': wide['单独销售金额'].fillna(0),
                    '搭售销售金额': wide['搭售销售金额'].fillna(0),
                    '适用主销品总销量': wide['适用主销品总销量'].fillna(0),
                    '被搭售的主销品yspu': wide['被搭售的主销品yspu'],
                    '被搭售的主销品yspu_code': wide[main_code_col],
                    '主销品单品销量': wide['大盘主销品单品销量'],
                    '主销品带动搭售量': wide['主销品带动搭售量'].fillna(0),
                    '场景': wide['场景'] if '场景' in wide.columns else '基础',
                    '占比': wide['占比'],
                    '整体搭售率': wide['整体搭售率'],
                    '明细搭售率': wide['明细搭售率']
                })
                
                st.session_state.wide_final = df_final
                st.success("🎉 宽表及衍生指标计算成功！大盘数据已完美通过代码自动对齐。")
                
            except Exception as e:
                st.error(f"❌ 数据清洗冲突，请核对字段格式。报错信息: {e}")

# ==========================================
# 3. 结果下载与动态线图可视化 Presentation
# ==========================================
if "wide_final" in st.session_state and st.session_state.wide_final is not None:
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
    
    yspu_list = sorted(df_res['搭售件名称'].dropna().unique().tolist())
    selected_yspu = st.selectbox("🎯 请选择或输入要分析的搭售件 (YSPU)：", yspu_list)
    
    if selected_yspu:
        df_chart = df_res[df_res['搭售件名称'] == selected_yspu].copy()
        df_chart['排序键'] = df_chart['销售月份'].astype(str) + "_" + df_chart['站点'].astype(str)
        df_chart = df_chart.sort_values(by='排序键')
        
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
        
        fig.add_trace(go.Scatter(
            x=df_chart['排序键'], y=df_chart['整体搭售率'],
            mode='lines+markers', name='整体搭售率',
            line=dict(color='#1f77b4', width=3), text=hover_texts, hoverinfo='text+y'
        ))
        
        fig.add_trace(go.Scatter(
            x=df_chart['排序键'], y=df_chart['明细搭售率'],
            mode='lines+markers', name='明细搭售率',
            line=dict(color='#ff7f0e', width=2, dash='dash'), text=hover_texts, hoverinfo='text+y'
        ))
        
        fig.update_layout(
            title=dict(text=f"📊 {selected_yspu} 的搭售率健康度走势分析（整体 vs 明细）", font=dict(size=18)),
            xaxis_title="时间轴与站点 (月份_国家)", yaxis_title="比率",
            yaxis=dict(tickformat=".2%"), hovermode="closest",
            legend=dict(orient="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            template="streamlit"
        )
        st.plotly_chart(fig, use_container_width=True)
else:
    st.info("💡 提示：请在左侧栏上传【1. IT搭售数据底表】和【2. 产品-场景套系映射表】，然后点击按钮启动清洗。")
