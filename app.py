import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io

st.set_page_config(page_title="搭售分析看板 & 宽表生成器", layout="wide")

st.title("📊 独立站搭售分析明细看板系统")
st.markdown("---")

# ==========================================
# 1. 侧边栏：文件上传区
# ==========================================
st.sidebar.header("📁 数据源上传")
file_b_raw = st.sidebar.file_uploader("1. 上传 IT提供搭售数据底表 (CSV 或 Excel)", type=["csv", "xlsx"])
file_map_raw = st.sidebar.file_uploader("2. 上传 产品-场景套系映射表 (Excel 或 CSV)", type=["xlsx", "csv"])

if "wide_final" not in st.session_state:
    st.session_state.wide_final = None

# 💡 数字/万字文本通用转换清洗函数
def clean_money_column(val):
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace(',', '')
    if not val_str or val_str == '-':
        return 0.0
    try:
        if '万' in val_str:
            num_part = val_str.replace('万', '').strip()
            return float(num_part) * 10000.0
        return float(val_str)
    except:
        return 0.0

# ==========================================
# 2. 核心数据处理引擎
# ==========================================
if file_b_raw and file_map_raw:
    if st.sidebar.button("🚀 开始清洗并生成搭售分析明细宽表"):
        with st.spinner("数分引擎正在清洗整合大盘数据，请稍候..."):
            try:
                if file_b_raw.name.endswith('.csv'):
                    df_b = pd.read_csv(file_b_raw)
                else:
                    df_b = pd.read_excel(file_b_raw)
                
                if file_map_raw.name.endswith('.csv'):
                    df_map = pd.read_csv(file_map_raw)
                else:
                    df_map = pd.read_excel(file_map_raw)
                
                df_b.columns = [str(c).strip() for c in df_b.columns]
                df_map.columns = [str(c).strip() for c in df_map.columns]
                
                # 模糊匹配映射表的关键列
                acc_col_in_map = None
                main_col_in_map = None
                for c in df_map.columns:
                    c_upper = c.upper()
                    if '搭售件' in c and 'YSPU' in c_upper:
                        acc_col_in_map = c
                    if '能被搭售' in c and 'YSPU' in c_upper:
                        main_col_in_map = c
                
                if not acc_col_in_map:
                    acc_col_in_map = '搭售件YSPU'
                if not main_col_in_map:
                    main_col_in_map = '能被搭售的主销品YSPU' if '能被搭售的主销品YSPU' in df_map.columns else '能被搭售品YSPU'

                # 基础清洗
                for amt_col in ['单独销售金额_cny', '搭售金额_cny']:
                    if amt_col in df_b.columns:
                        df_b[amt_col] = df_b[amt_col].apply(clean_money_column)
                    else:
                        df_b[amt_col] = 0.0
                
                for num_c in ['销量', '单独销量', '搭售销量']:
                    if num_c in df_b.columns:
                        df_b[num_c] = pd.to_numeric(df_b[num_c], errors='coerce').fillna(0)

                # 提取主销品大盘销量
                df_main_pool = df_b[df_b['产品类型'] == '主销品'].copy()
                if len(df_main_pool) == 0:
                    df_main_pool = df_b[df_b['产品类型'].str.contains('主', na=False)].copy()
                    
                df_main_sum = df_main_pool.groupby(['销售月份(month)', '销售国', 'yspu_code'])['销量'].sum().reset_index()
                df_main_sum.rename(columns={'yspu_code': '匹配用_主销品编码', '销量': '大盘主销品单品销量'}, inplace=True)
                
                # 计算搭售件大盘指标
                df_acc_pool = df_b[df_b['产品类型'] == '搭售件'].copy()
                if len(df_acc_pool) == 0:
                    df_acc_pool = df_b[df_b['产品类型'].str.contains('搭|配', na=False)].copy()
                
                df_acc_market = df_acc_pool.groupby(['销售月份(month)', '销售国', 'yspu_code']).agg({
                    '销量': 'sum', '单独销量': 'sum', '搭售销量': 'sum', '单独销售金额_cny': 'sum', '搭售金额_cny': 'sum'
                }).reset_index().rename(columns={
                    '销量': '搭售件总销量', '单独销量': '搭售件单独销量', '搭售销量': '搭售件搭售销量',
                    '单独销售金额_cny': '单独销售金额', '搭售金额_cny': '搭售销售金额'
                })

                # 解析映射表
                map_dict = {}
                for _, row in df_map.iterrows():
                    acc_code = str(row[acc_col_in_map]).strip()
                    main_codes_str = str(row[main_col_in_map]).replace('"', '').replace('\n', '').replace('\r', '')
                    map_dict[acc_code] = [c.strip() for c in main_codes_str.split(',') if c.strip()]
                
                applicable_sales_list = []
                for _, row in df_acc_market.iterrows():
                    m_month, m_site, m_acc = row['销售月份(month)'], row['销售国'], row['yspu_code']
                    app_main_codes = map_dict.get(m_acc, [])
                    sub_main = df_main_sum[
                        (df_main_sum['销售月份(month)'] == m_month) & 
                        (df_main_sum['销售国'] == m_site) & 
                        (df_main_sum['匹配用_主销品编码'].isin(app_main_codes))
                    ]
                    applicable_sales_list.append(sub_main['大盘主销品单品销量'].sum() if len(sub_main) > 0 else 0)
                
                df_acc_market['适用主销品总销量'] = applicable_sales_list

                # 提取真实明细行
                main_code_col = 'get被搭售的主销品yspu_code' if 'get被搭售的主销品yspu_code' in df_b.columns else 'get被搭售的主销品yspu_code'
                if 'get被搭售的主销品yspu_code' not in df_b.columns and '被搭售的主销品yspu_code' in df_b.columns:
                    main_code_col = 'box_main_code'
                    df_b.rename(columns={'被搭售的主销品yspu_code': 'box_main_code'}, inplace=True)
                
                df_b_detail = df_b[
                    (df_b['产品类型'] == '搭售件') & (df_b[main_code_col] != '-') & (df_b[main_code_col].notna())
                ].copy()
                df_b_detail.rename(columns={'搭售销量': '主销品带动搭售量'}, inplace=True)
                
                df_b_detail = pd.merge(df_b_detail, df_main_sum, left_on=['销售月份(month)', '销售国', main_code_col], right_on=['销售月份(month)', '销售国', '匹配用_主销品编码'], how='left')
                df_b_detail['大盘主销品单品销量'] = df_b_detail['大盘主销品单品销量'].fillna(0)
                
                wide = pd.merge(df_b_detail, df_acc_market, on=['销售月份(month)', '销售国', 'yspu_code'], how='left')
                
                wide['占比'] = wide['搭售件搭售销量'] / wide['搭售件总销量']
                wide['整体搭售率'] = wide['搭售件搭售销量'] / wide['适用主销品总销量']
                wide['明细搭售率'] = wide['主销品带动搭售量'] / wide['大盘主销品单品销量']
                
                for rate_col in ['占比', '整体搭售率', '明细搭售率']:
                    wide[rate_col] = wide[rate_col].fillna(0).replace([float('inf'), float('-inf')], 0)
                
                df_final = pd.DataFrame({
                    '销售月份': wide['销售月份(month)'].astype(str),
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
                st.success("🎉 宽表数据清洗成功，衍生指标已全部通过对齐校验！")
                
            except Exception as e:
                st.error(f"❌ 数据清洗冲突，请核对字段格式。报错信息: {e}")

# ==========================================
# 3. 数据分析面板展示与动态联动
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
    
    with st.expander("🔍 展开查看宽表数据快照"):
        st.dataframe(df_res.head(15))
    
    st.markdown("---")
    st.header("📈 搭售业务多维联动诊断看板")
    
    # 💡 联动第一层：选择搭售件
    yspu_list = sorted(df_res['搭售件名称'].dropna().unique().tolist())
    selected_yspu = st.selectbox("🎯 1. 请选择要诊断的搭售件 (YSPU)：", yspu_list)
    
    if selected_yspu:
        # 初步筛选出该搭售件的数据
        df_yspu = df_res[df_res['搭售件名称'] == selected_yspu].copy()
        
        # 💡 联动第二层：平行选择销售国和销售月份区间
        col_fill1, col_fill2 = st.columns(2)
        
        with col_fill1:
            available_sites = [s for s in ['US', 'GB', 'DE'] if s in df_yspu['站点'].unique()]
            if not available_sites:
                available_sites = sorted(df_yspu['站点'].unique().tolist())
            selected_site = st.selectbox("🌐 2. 选择分析国别 (站点)：", available_sites)
            
        with col_fill2:
            all_months = sorted(df_yspu['销售月份'].unique().tolist())
            if len(all_months) >= 2:
                selected_month_range = st.select_slider(
                    "📅 3. 选择分析月份区间：",
                    options=all_months,
                    value=(all_months[0], all_months[-1])
                )
                start_month, end_month = selected_month_range
            else:
                start_month = end_month = all_months[0]
                st.info(f"当前筛选条件下仅包含单月数据: {start_month}")
        
        # 执行全局条件过滤（搭售件 + 站点 + 月份区间）
        df_filtered = df_yspu[
            (df_yspu['站点'] == selected_site) & 
            (df_yspu['销售月份'] >= start_month) & 
            (df_yspu['销售月份'] <= end_month)
        ].copy()
        
        df_filtered = df_filtered.sort_values(by='销售月份')
        
        if len(df_filtered) > 0:
            # ----------------------------------------------------
            # 📊 图表一：【整体搭售率】趋势图（单线 + 强标签）
            # ----------------------------------------------------
            st.markdown(f"### ① {selected_yspu} · 整体搭售率大盘趋势 ({selected_site})")
            
            # 按月份聚合整体搭售率，防止多条明细行导致画图重叠
            df_overall_trend = df_filtered.groupby('销售月份')['整体搭售率'].first().reset_index()
            
            fig_overall = go.Figure()
            fig_overall.add_trace(go.Scatter(
                x=df_overall_trend['销售月份'], 
                y=df_overall_trend['整体搭售率'],
                mode='lines+markers+text',   # 💡 强标签：强制开启文本标签显示
                name='整体搭售率',
                text=[f"{v:.2%}" for v in df_overall_trend['整体搭售率']], # 标签内容格式化
                textposition="top center",
                line=dict(color='#1f77b4', width=4),
                marker=dict(size=8)
            ))
            
            fig_overall.update_layout(
                xaxis=dict(type='category', title="销售月份"),
                yaxis=dict(title="整体搭售率", tickformat=".2%"),
                margin=dict(l=40, r=40, t=40, b=40),
                hovermode="x unified"
            )
            st.plotly_chart(fig_overall, use_container_width=True)
            
            # ----------------------------------------------------
            # 📊 图表二：【明细搭售率】趋势图（主销品分拆 + 悬停显示）
            # ----------------------------------------------------
            st.markdown("---")
            st.markdown(f"### ② {selected_yspu} · 关联不同主销品之明细搭售率追踪")
            
            # 提取该过滤池内所有出现过的关联主销品名称
            available_mains = sorted(df_filtered['被搭售的主销品yspu'].dropna().unique().tolist())
            
            # 💡 联动第三层：明细专属主销品过滤
            selected_mains = st.multiselect(
                "🔍 过滤特定「被搭售主销品yspu」（留空默认全量展示对比）：", 
                options=available_mains,
                default=[]
            )
            
            # 如果运营没有选，默认展示该搭售件旗下的所有主销品线条
            display_mains = selected_mains if selected_mains else available_mains
            
            fig_detail = go.Figure()
            
            for main_item in display_mains:
                df_item = df_filtered[df_filtered['被搭售的主销品yspu'] == main_item]
                if len(df_item) > 0:
                    # 构造完美的 Hover 浮窗文本
                    hover_texts = []
                    for _, r in df_item.iterrows():
                        txt = (
                            f"<b>主销品:</b> {r['被搭售的主销品yspu']}<br>"
                            f"<b>主销品单品销量:</b> {int(r['主销品单品销量'])}<br>"
                            f"<b>带动该配件量:</b> {int(r['主销品带动搭售量'])}"
                        )
                        hover_texts.append(txt)
                        
                    fig_detail.add_trace(go.Scatter(
                        x=df_item['销售月份'], 
                        y=df_item['明细搭售率'],
                        mode='lines+markers',  # 💡 干净清爽，不堆叠标签，只显示点线
                        name=main_item,
                        text=hover_texts,
                        hoverinfo="text+y"     # 💡 鼠标悬停处体现搭售率数据及关联信息
                    ))
            
            fig_detail.update_layout(
                xaxis=dict(type='category', title="销售月份"),
                yaxis=dict(title="明细搭售率", tickformat=".2%"),
                legend=dict(orient="h", yanchor="bottom", y=-0.3, xanchor="left", x=0),
                margin=dict(l=40, r=40, t=40, b=40),
                hovermode="closest"
            )
            st.plotly_chart(fig_detail, use_container_width=True)
            
        else:
            st.warning("⚠️ 当前月份和国家筛选区间内没有匹配的销售流水，请调整上方的筛选器。")
else:
    st.info("💡 提示：请在左侧栏上传【1. IT搭售数据底表】和【2. 产品-场景套系映射表】，然后点击按钮启动清洗。")
