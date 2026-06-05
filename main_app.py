import streamlit as st
import pandas as pd
import io

from fabric_cost import run_fabric_calculation
from shipping_cost import run_shipping_calculation
from return_analyzer import run_return_analysis
from keyword_analyzer import run_keyword_analysis
from visualizer import create_keyword_trend_fig, prepare_color_sales_data, create_color_sales_fig

st.set_page_config(page_title="公司数据处理平台", layout="wide")
st.title("📊 自动化计算工具平台")

st.sidebar.title("工具菜单")
menu_choice = st.sidebar.radio(
    "请选择你需要使用的功能：",
    ["🧵 布料费用计算", "🚚 运费计算", "📦 退货数据分析", "📊 数据可视化", "✂️ 排料计算 (开发中)"]
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
