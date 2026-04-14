import streamlit as st
import pandas as pd
import io

# 导入你的三位“大厨”
from fabric_cost import run_fabric_calculation
from shipping_cost import run_shipping_calculation
from return_analyzer import run_return_analysis

# 设置网页
st.set_page_config(page_title="公司数据处理平台", layout="wide")
st.title("📊 自动化计算工具平台")

# 左侧菜单栏
st.sidebar.title("工具菜单")
menu_choice = st.sidebar.radio(
    "请选择你需要使用的功能：",
    ["🧵 布料费用计算", "🚚 运费计算", "📦 退货数据分析", "✂️ 排料计算 (开发中)"]
)

# ----------------------------------------
# 工具1：布料费用计算
# ----------------------------------------
if menu_choice == "🧵 布料费用计算":
    st.header("布料费用自动化计算")
    st.write("请上传包含【布料跟踪表, 订单表, SKU-款号映射, 款号信息表】的 Excel 文件")
    uploaded_file = st.file_uploader("点击此处上传 Excel (布料)", type=["xlsx", "xls"])

    if uploaded_file is not None:
        if st.button("🚀 开始计算"):
            with st.spinner("程序正在疯狂计算中，包含两轮备布数据处理..."):
                try:
                    result_dict = run_fabric_calculation(uploaded_file)
                    df_results = result_dict["布料费用结果"]
                    df_backup = result_dict["备布数据"]

                    if not df_results.empty:
                        st.success(f"🎉 计算成功！共计算了 {len(df_results)} 个款号，发现 {len(df_backup)} 条备布记录。")

                        st.write("👀 预览前 10 条主数据结果：")
                        st.dataframe(df_results.head(10))

                        # 把两个表格写进同一个 Excel 文件的不同 Sheet 里
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_results.to_excel(writer, sheet_name='布料费用结果', index=False)
                            if not df_backup.empty:
                                df_backup.to_excel(writer, sheet_name='备布数据', index=False)
                        output.seek(0)

                        st.download_button(
                            label="📥 下载完整计算结果 (含备布数据)",
                            data=output,
                            file_name="布料费用计算结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("⚠️ 计算完成，但没有提取到符合条件的有效数据。")
                except Exception as e:
                    st.error(f"❌ 计算过程中发生错误：{str(e)}")

# ----------------------------------------
# 工具2：运费计算
# ----------------------------------------
elif menu_choice == "🚚 运费计算":
    st.header("运费自动化计算")
    st.write("请上传包含【发货日报表, FBA头程, SKU-款号映射, 款号信息表】的 Excel 文件")
    uploaded_file = st.file_uploader("点击此处上传 Excel (运费)", type=["xlsx", "xls"])
    if uploaded_file is not None:
        if st.button("🚀 开始计算"):
            with st.spinner("程序正在疯狂计算运费中，请稍候..."):
                try:
                    result_dict = run_shipping_calculation(uploaded_file)
                    if result_dict is not None:
                        st.success("🎉 运费计算成功！")
                        st.write("👀 预览前 10 条运费结果：")
                        st.dataframe(result_dict["运费结果"].head(10))

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_dict["运费结果"].to_excel(writer, sheet_name='运费结果', index=False)
                            result_dict["df4明细数据"].to_excel(writer, sheet_name='df4明细数据', index=False)
                        output.seek(0)

                        st.download_button("📥 下载完整计算结果 (含明细)", data=output, file_name="运费计算结果.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    else:
                        st.warning("⚠️ 计算完成，但没有提取到有效的运费数据。")
                except Exception as e:
                    st.error(f"❌ 计算过程中发生错误：{str(e)}")

# ----------------------------------------
# 工具3：退货数据分析
# ----------------------------------------
elif menu_choice == "📦 退货数据分析":
    st.header("退货数据自动化分析")
    st.info("💡 提示：按住键盘的 Ctrl (或 Command) 键，可以一次性选中多个 txt 文件上传。")

    col1, col2, col3 = st.columns(3)
    with col1:
        order_files = st.file_uploader("1️⃣ 上传 order 销售数据 (.txt)", type=["txt"], accept_multiple_files=True)
    with col2:
        return_files = st.file_uploader("2️⃣ 上传 return 退货数据 (.txt)", type=["txt"], accept_multiple_files=True)
    with col3:
        sku_file = st.file_uploader("3️⃣ 上传 SKU信息表 (.xls/.xlsx)", type=["xls", "xlsx"])

    if order_files and return_files and sku_file:
        if st.button("🚀 开始合并与分析"):
            with st.spinner("正在处理海量数据，请耐心等待..."):
                try:
                    final_df = run_return_analysis(order_files, return_files, sku_file)
                    st.success(f"🎉 处理完成！共生成 {len(final_df)} 条源数据记录。")

                    st.write("👀 数据预览：")
                    st.dataframe(final_df.head(20))

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, sheet_name='源数据', index=False)
                    output.seek(0)

                    st.download_button("📥 下载退货分析源数据 (Excel)", data=output, file_name="退货分析源数据.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"❌ 处理过程中发生错误：{str(e)}")

# ----------------------------------------
# 工具4：排料计算
# ----------------------------------------
elif menu_choice == "✂️ 排料计算 (开发中)":
    st.info("排料计算功能工程师正在熬夜开发中，敬请期待...")