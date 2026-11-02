import streamlit as st
import pandas as pd
import io

from fabric_cost import run_fabric_calculation
from shipping_cost import run_shipping_calculation
from return_analyzer import run_return_analysis
from keyword_analyzer import run_keyword_analysis
from visualizer import create_keyword_trend_fig, prepare_color_sales_data, create_color_sales_fig, create_color_sales_pie_fig

# --- 新增模块：采购单自动化生成（v1.5）---
from purchase_order_generator.ui import render_purchase_order_page

# --- 新增模块：成本计算（新）--- 新输入口径（装箱明细 + 款号档案出货数），与上面两个旧模块完全独立
from cost_new import run_cost_calculation, export_excel

st.set_page_config(page_title="公司数据处理平台", layout="wide")
st.title("📊 自动化计算工具平台")

st.sidebar.title("工具菜单")
menu_choice = st.sidebar.radio(
    "请选择你需要使用的功能：",
    ["🧵 布料费用计算", "🚚 运费计算", "🧮 成本计算（新）",
     "📦 退货数据分析", "📊 数据可视化", "✂️ 排料计算 (开发中)", "📋 采购单自动化生成"]
)

if menu_choice == "🧵 布料费用计算":
    st.header("布料费用自动化计算")
    uploaded_file = st.file_uploader("点击此处上传 Excel (布料)", type=["xlsx", "xls"])
    if uploaded_file is not None:
        if st.button("🚀 开始计算"):
            with st.spinner("计算中..."):
                try:
                    result_dict = run_fabric_calculation(uploaded_file)
                    df_results = result_dict["布料费用结果"]
                    df_backup = result_dict["备布数据"]
                    if not df_results.empty:
                        st.success("🎉 计算成功！")
                        st.dataframe(df_results.head(10))
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_results.to_excel(writer, sheet_name='布料费用结果', index=False)
                            if not df_backup.empty:
                                df_backup.to_excel(writer, sheet_name='备布数据', index=False)
                        output.seek(0)
                        st.download_button(
                            "📥 下载完整结果",
                            data=output,
                            file_name="布料结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"❌ 错误：{str(e)}")

elif menu_choice == "🚚 运费计算":
    st.header("运费自动化计算")
    uploaded_file = st.file_uploader("点击此处上传 Excel (运费)", type=["xlsx", "xls"])
    if uploaded_file is not None:
        if st.button("🚀 开始计算"):
            with st.spinner("计算中..."):
                try:
                    result_dict = run_shipping_calculation(uploaded_file)
                    if result_dict is not None:
                        st.success("🎉 计算成功！")
                        st.dataframe(result_dict["运费结果"].head(10))
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_dict["运费结果"].to_excel(writer, sheet_name='运费结果', index=False)
                            result_dict["df4明细数据"].to_excel(writer, sheet_name='df4明细数据', index=False)
                        output.seek(0)
                        st.download_button(
                            "📥 下载完整结果",
                            data=output,
                            file_name="运费结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"❌ 错误：{str(e)}")


elif menu_choice == "🧮 成本计算（新）":
    st.header("成本计算（新）")
    st.caption(
        "**新口径**：一个 Excel 一次算出布料费 + 运费。与「布料费用计算」「运费计算」两个旧模块互不影响——"
        "旧模块走订单表 / 发货日报表，本模块走装箱明细，成衣数直接读款号档案。"
    )

    with st.expander("📋 需要的 4 张工作表", expanded=False):
        st.markdown(
            """
| 工作表 | 必需列 | 说明 |
|---|---|---|
| **布料跟踪表** | 下单日期 / 订单编号 / 产品名称 / 颜色&色号 / 布料送货数量（匹） / 金额（元） | 匹数或金额为空的「未送货」行自动剔除 |
| **款号档案** | 款号 / **出货数（件/匹）** | 这一列就是成衣数 Q；**留空 = 该款已淘汰，不参与计算** |
| **FBA头程** | 发货日期 / 货件编号 / 发货数量pcs / 总金额 | 头程费用来源 |
| **数据源-领星-FBA货件-装箱明细** | 货件单号 / MSKU / 申报量 | 按各 SKU 申报量占比分摊运费 |
            """
        )
        st.info("时间窗口统一取近一年（365 天）：布料按最新下单日、运费按最新发货日各自倒推。"
                "装箱明细的货件单号是合并单元格，程序会自动向下填充。")

    uploaded_file = st.file_uploader(
        "点击此处上传 Excel（一个文件即可）",
        type=["xlsx", "xls"],
        key="cost_new_uploader"
    )

    if uploaded_file is not None:
        if st.button("🚀 开始计算", key="cost_new_run"):
            with st.spinner("正在计算布料与运费..."):
                try:
                    uploaded_file.seek(0)
                    result = run_cost_calculation(uploaded_file)
                    summary = result.get("摘要", {})
                    df_main = result["成本汇总"]

                    if df_main.empty:
                        st.warning(
                            "没有算出任何款。请检查：①「款号档案 - 出货数（件/匹）」是否已填写；"
                            "②「布料跟踪表」在近一年内是否有已送货的匹数与金额。")
                    else:
                        st.success(
                            f"🎉 完成！{summary.get('完整成本款数', 0)} 个款拿到完整成本"
                            f"（布料 {summary.get('布料款数', 0)} 个 · 运费 {summary.get('运费款数', 0)} 个）")

                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("单件总成本",
                                  f"¥{summary.get('合计最低', 0):.2f} ~ {summary.get('合计最高', 0):.2f}")
                        c2.metric("运费金额覆盖率", str(summary.get("运费覆盖率", "-")))
                        c3.metric("发货件数合计", f"{summary.get('发货件数合计', 0):,}")
                        c4.metric("未进主表的款", summary.get("排除款数", 0))

                        st.markdown("### 💰 成本汇总")
                        st.dataframe(df_main, use_container_width=True)

                        df_only = result.get("仅有运费的款")
                        if df_only is not None and not df_only.empty:
                            with st.expander(f"ℹ️ 另有 {len(df_only)} 个款只有运费、没有布料成本"):
                                st.caption("这些款有发货记录，但近一年没有布料送货、或已被淘汰。")
                                st.dataframe(df_only, use_container_width=True)

                        df_ex = result.get("排除款清单")
                        if df_ex is not None and not df_ex.empty:
                            with st.expander(f"ℹ️ 已排除 {len(df_ex)} 个款（不参与计算）"):
                                st.dataframe(df_ex, use_container_width=True)

                        buf = io.BytesIO()
                        export_excel(result, buf)
                        buf.seek(0)
                        st.download_button(
                            "📥 下载完整结果（8 个 sheet）",
                            data=buf,
                            file_name="成本计算结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="cost_new_download"
                        )
                except Exception as e:
                    st.error(f"❌ 错误：{str(e)}")
                    with st.expander("查看详细错误"):
                        st.exception(e)

elif menu_choice == "📦 退货数据分析":
    st.header("退货数据自动化分析")
    col1, col2, col3 = st.columns(3)
    with col1:
        order_files = st.file_uploader("1️⃣ order 数据 (.txt)", type=["txt"], accept_multiple_files=True)
    with col2:
        return_files = st.file_uploader("2️⃣ return 数据 (.txt)", type=["txt"], accept_multiple_files=True)
    with col3:
        sku_file = st.file_uploader("3️⃣ SKU信息表 (.xls/.xlsx)", type=["xls", "xlsx"])

    if order_files and return_files and sku_file:
        if st.button("🚀 开始分析"):
            with st.spinner("处理中..."):
                try:
                    final_df = run_return_analysis(order_files, return_files, sku_file)
                    st.success("🎉 处理完成！")
                    st.dataframe(final_df.head(10))
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, sheet_name='源数据', index=False)
                    output.seek(0)
                    st.download_button(
                        "📥 下载结果",
                        data=output,
                        file_name="退货源数据.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"❌ 错误：{str(e)}")
elif menu_choice == "📊 数据可视化":
    st.header("数据可视化")
    visualization_choice = st.sidebar.radio(
        "请选择可视化类型：",
        ["流量占比可视化", "颜色销量可视化"]
    )

    if visualization_choice == "流量占比可视化":
        st.subheader("西柚搜索词流量趋势可视化")
        st.write("请上传西柚搜索词的 Excel 文件，支持命名格式如 `202501.xlsx` 或 `2025-01.xlsx`")
        st.markdown("### ⚙️ 仪表盘设置")
        col_param1, col_param2 = st.columns(2)
        with col_param1:
            threshold_Rank = st.slider(
                "🏆 提取每个月流量占比前 N 的关键词",
                min_value=1,
                max_value=100,
                value=30,
                step=1
            )
        with col_param2:
            maxNum_horizontal = st.slider(
                "🪟 每行显示的图表数量",
                min_value=1,
                max_value=4,
                value=2,
                step=1
            )
        uploaded_files = st.file_uploader(
            "点击此处批量上传 Excel 数据表",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="keyword_visualizer_files"
        )
        if uploaded_files:
            if st.button("🚀 生成数据大屏", key="generate_keyword_visualizer"):
                with st.spinner("正在清洗数据并绘制图表..."):
                    try:
                        df2 = run_keyword_analysis(uploaded_files, threshold_Rank)
                        st.success(f"🎉 成功提取每个月流量占比前 {threshold_Rank} 的关键词并生成分析数据！")
                        st.markdown("### 📊 关键词流量趋势折线图")
                        # 1. 聚合计算每个关键词的 流量占比 总和作为排序依据
                        kw_rank = (
                            df2.groupby('关键词 (数据来源于西柚找词)')['流量占比']
                            .sum()
                            .sort_values(ascending=False)
                        )
                        # 2. 按流量占比倒序取关键词，不含空值
                        keywords = kw_rank.index.dropna().astype(str).tolist()
                        cols = st.columns(maxNum_horizontal)
                        # 3. 按关键词排序循环绘图
                        for i, kw in enumerate(keywords):
                            col = cols[i % maxNum_horizontal]
                            kw_data = df2[df2['关键词 (数据来源于西柚找词)'].astype(str) == kw].copy()
                            fig = create_keyword_trend_fig(kw_data, f"#{i+1} {kw}")
                            col.plotly_chart(fig, use_container_width=True)
                        st.markdown("### 💾 源数据下载")
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df2.to_excel(writer, sheet_name='Top关键词趋势数据', index=False)
                        output.seek(0)
                        st.download_button(
                            "📥 下载清洗后的数据表",
                            data=output,
                            file_name="搜索词分析结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"❌ 分析过程中发生错误：{str(e)}")

    elif visualization_choice == "颜色销量可视化":
        st.subheader("竞品颜色销量可视化")
        uploaded_file = st.file_uploader(
            "点击此处上传竞品信息 Excel 数据表",
            type=["xlsx", "xls"],
            key="color_sales_visualizer_file"
        )

        if uploaded_file is not None:
            try:
                color_sales_df, color_sales_long_df = prepare_color_sales_data(uploaded_file)
                color_count = len(color_sales_df)

                if color_count == 0:
                    st.warning("没有找到可用于展示的颜色销量数据")
                else:
                    st.markdown("### ⚙️ 仪表盘设置")
                    col_param1, col_param2 = st.columns(2)
                    with col_param1:
                        chart_count = st.slider(
                            "📈 图表个数",
                            min_value=1,
                            max_value=color_count,
                            value=min(6, color_count),
                            step=1
                        )
                    with col_param2:
                        maxNum_horizontal = st.slider(
                            "🪟 每行可展示的最大图表数量",
                            min_value=1,
                            max_value=3,
                            value=2,
                            step=1
                        )

                    # ===== 预览板块：饼状图 + 销量排名表 =====
                    st.markdown("### 👀 数据预览")

                    preview_col1, preview_col2 = st.columns(2)
                    with preview_col1:
                        pie_fig = create_color_sales_pie_fig(color_sales_df)
                        preview_col1.plotly_chart(pie_fig, use_container_width=True)

                    with preview_col2:
                        ranking_df = (
                            color_sales_df[['颜色名称', '颜色总销量']]
                            .head(chart_count)
                            .reset_index(drop=True)
                        )
                        ranking_df.index = ranking_df.index + 1  # 序号从 1 开始
                        ranking_df = ranking_df.rename(columns={'颜色总销量': '颜色销量'})
                        st.dataframe(
                            ranking_df,
                            use_container_width=True,
                            column_config={
                                "颜色名称": "颜色名称",
                                "颜色销量": st.column_config.NumberColumn("颜色销量", format="%d")
                            }
                        )

                    if st.button("🚀 生成颜色销量图表", key="generate_color_sales_visualizer"):
                        st.success(f"🎉 成功识别 {color_count} 个颜色，已按颜色总销量倒序展示前 {chart_count} 个")
                        st.markdown("### 📊 颜色销量趋势折线图")

                        selected_colors = color_sales_df.head(chart_count)
                        cols = st.columns(maxNum_horizontal)
                        for i, row in selected_colors.iterrows():
                            color_name = row['颜色名称']
                            total_sales = row['颜色总销量']
                            color_data = color_sales_long_df[color_sales_long_df['颜色名称'] == color_name].copy()
                            fig = create_color_sales_fig(color_data, color_name, total_sales)
                            cols[i % maxNum_horizontal].plotly_chart(fig, use_container_width=True)

                        st.markdown("### 💾 源数据下载")
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            color_sales_df.to_excel(writer, sheet_name='颜色销量汇总', index=False)
                            color_sales_long_df.to_excel(writer, sheet_name='颜色销量趋势明细', index=False)
                        output.seek(0)
                        st.download_button(
                            "📥 下载清洗后的数据表",
                            data=output,
                            file_name="竞品颜色销量分析结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            except Exception as e:
                st.error(f"❌ 分析过程中发生错误：{str(e)}")


elif menu_choice == "✂️ 排料计算 (开发中)":
    st.info("排料计算功能工程师正在熬夜开发中，敬请期待...")


# --- 新增模块：采购单自动化生成 ---
elif menu_choice == "📋 采购单自动化生成":
    st.header("采购单自动化生成")
    render_purchase_order_page()