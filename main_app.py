import streamlit as st
import pandas as pd
import io
import plotly.express as px

# 导入你的四位“大厨”
from fabric_cost import run_fabric_calculation
from shipping_cost import run_shipping_calculation
from return_analyzer import run_return_analysis
from keyword_analyzer import run_keyword_analysis, get_top_keywords
from plotting_utils import plot_keyword_trends

# 设置网页（开启 wide 宽屏模式，方便展示多个图表）
st.set_page_config(page_title="公司数据处理平台", layout="wide")
st.title("📊 自动化计算工具平台")

# 左侧菜单栏
st.sidebar.title("工具菜单")
menu_choice = st.sidebar.radio(
    "请选择你需要使用的功能：",
    ["🧵 布料费用计算", "🚚 运费计算", "📦 退货数据分析", "📈 搜索词趋势分析", "✂️ 排料计算 (开发中)"]
)

# ----------------------------------------
# 工具1：布料费用计算
# ----------------------------------------
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
                        st.download_button("📥 下载完整结果", data=output, file_name="布料结果.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"❌ 错误：{str(e)}")

# ----------------------------------------
# 工具2：运费计算
# ----------------------------------------
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
                        st.download_button("📥 下载完整结果", data=output, file_name="运费结果.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"❌ 错误：{str(e)}")

# ----------------------------------------
# 工具3：退货数据分析
# ----------------------------------------
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
                    st.download_button("📥 下载结果", data=output, file_name="退货源数据.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"❌ 错误：{str(e)}")

# ----------------------------------------
# 工具4：搜索词趋势分析 (带数据可视化)
# ----------------------------------------
elif menu_choice == "📈 搜索词趋势分析":
    st.header("西柚搜索词流量趋势可视化")
    Num_Rank = st.slider("🏆 关键词排名数", 5, 100, 30)
    maxNum_horizontal = st.slider("🪟 每行图表数", 1, 4, 2)
    uploaded_files = st.file_uploader("批量上传 Excel", type=["xlsx"], accept_multiple_files=True)
    if uploaded_files:
        if st.button("🚀 生成趋势图"):
            with st.spinner("绘图中..."):
                try:
                    df1 = run_keyword_analysis(uploaded_files, Num_Rank)
                    st.success(f"成功提取排名前 {Num_Rank} 关键词数据")
                    keywords = get_top_keywords(df1)
                    cols = st.columns(maxNum_horizontal)
                    for i, kw in enumerate(keywords):
                        col = cols[i % maxNum_horizontal]
                        kw_data = df1[df1['关键词 (数据来源于西柚找词)'] == kw].copy()
                        fig = plot_keyword_trends(kw_data, kw)
                        col.plotly_chart(fig, use_container_width=True)
                    # 下载数据
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df1.to_excel(writer, sheet_name='Top关键词趋势数据', index=False)
                    output.seek(0)
                    st.download_button("📥 下载分析数据", data=output, file_name="搜索词分析结果.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"绘图失败：{str(e)}")

# ----------------------------------------
# 工具5：排料计算
# ----------------------------------------
elif menu_choice == "✂️ 排料计算 (开发中)":
    st.info("排料计算功能工程师正在熬夜开发中，敬请期待...")