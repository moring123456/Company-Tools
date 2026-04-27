import streamlit as st
import pandas as pd
import io
import plotly.express as px

# 导入你的四位“大厨”
from fabric_cost import run_fabric_calculation
from shipping_cost import run_shipping_calculation
from return_analyzer import run_return_analysis
from keyword_analyzer import run_keyword_analysis

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
    st.write("请上传西柚搜索词的 Excel 文件，文件名必须为 `YYYY-MM.xlsx` 格式（如 2025-01.xlsx）")

    # 控制参数区
    st.markdown("### ⚙️ 仪表盘设置")
    col_param1, col_param2 = st.columns(2)
    with col_param1:
        Num_Rank = st.slider("🏆 提取流量排名前 N 的关键词 (Num_Rank)", min_value=5, max_value=100, value=30, step=1)
    with col_param2:
        maxNum_horizontal = st.slider("🪟 每行显示的图表数量 (maxNum_horizontal)", min_value=1, max_value=4, value=2,
                                      step=1)

    # 文件上传区
    uploaded_files = st.file_uploader("点击此处批量上传 Excel 数据表", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files:
        if st.button("🚀 生成数据大屏"):
            with st.spinner("正在清洗数据并绘制图表..."):
                try:
                    # 获取清洗后的数据 df1
                    df1 = run_keyword_analysis(uploaded_files, Num_Rank)
                    st.success(f"🎉 成功提取排名前 {Num_Rank} 的关键词！")

                    # ---- 开始绘制图表 ----
                    st.markdown("### 📊 关键词流量趋势折线图")

                    # 获取去重并按排名排序好的关键词列表
                    keywords = df1['关键词 (数据来源于西柚找词)'].unique()

                    # 动态创建指定数量的列
                    cols = st.columns(maxNum_horizontal)

                    # 遍历每一个关键词，分配到对应的列中
                    for i, kw in enumerate(keywords):
                        col = cols[i % maxNum_horizontal]  # 取余数，实现循环分配到每一列

                        # 过滤当前关键词的数据
                        kw_data = df1[df1['关键词 (数据来源于西柚找词)'] == kw].copy()
                        # 确保月份排序正确 (01, 02, 03...)
                        kw_data = kw_data.sort_values('月份')

                        # 使用 Plotly 绘制多线图
                        fig = px.line(
                            kw_data,
                            x="月份",
                            y="流量占比",
                            color="年份",  # 颜色按年份区分（有几年画几条线）
                            title=f"#{i + 1} {kw}",
                            markers=True  # 显示折线上的数据点
                        )

                        # 优化图表显示
                        fig.update_layout(
                            xaxis_title="月份",
                            yaxis_title="流量占比",
                            yaxis_tickformat='.2%',  # Y轴显示为百分比
                            xaxis_type='category',  # 确保月份被当作类别而不是连续数字
                            margin=dict(l=20, r=20, t=40, b=20),
                            height=350
                        )

                        # 在对应的列展示图表
                        col.plotly_chart(fig, use_container_width=True)

                    # 提供基础数据下载
                    st.markdown("### 💾 源数据下载")
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df1.to_excel(writer, sheet_name='Top关键词趋势数据', index=False)
                    output.seek(0)
                    st.download_button("📥 下载清洗后的数据表", data=output, file_name="搜索词分析结果.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                except Exception as e:
                    st.error(f"❌ 分析过程中发生错误：{str(e)}")

# ----------------------------------------
# 工具5：排料计算
# ----------------------------------------
elif menu_choice == "✂️ 排料计算 (开发中)":
    st.info("排料计算功能工程师正在熬夜开发中，敬请期待...")