import streamlit as st
import pandas as pd
import io

# 导入你的两位“大厨”
from fabric_cost import run_fabric_calculation
from shipping_cost import run_shipping_calculation

# 设置网页
st.set_page_config(page_title="公司数据处理平台", layout="wide")
st.title("📊 自动化计算工具平台")

# 左侧菜单栏
st.sidebar.title("工具菜单")
menu_choice = st.sidebar.radio(
    "请选择你需要使用的功能：",
    ["🧵 布料费用计算", "🚚 运费计算", "✂️ 排料计算 (开发中)"]  # 运费计算上线啦！
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
            with st.spinner("程序正在疯狂计算中..."):
                try:
                    result_df = run_fabric_calculation(uploaded_file)
                    if result_df is not None and not result_df.empty:
                        st.success("🎉 布料费用计算成功！")
                        st.dataframe(result_df)

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_df.to_excel(writer, index=False)
                        output.seek(0)

                        st.download_button(
                            label="📥 下载计算结果 (Excel格式)",
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
                        st.success("🎉 运费计算成功！包含【运费结果】和【明细数据】。")

                        # 网页上只预览主结果，不然太占地方
                        st.write("预览前 10 条运费结果：")
                        st.dataframe(result_dict["运费结果"].head(10))

                        # 把两个表格写进同一个 Excel 文件的不同 Sheet 里
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_dict["运费结果"].to_excel(writer, sheet_name='运费结果', index=False)
                            result_dict["df4明细数据"].to_excel(writer, sheet_name='df4明细数据', index=False)
                        output.seek(0)

                        st.download_button(
                            label="📥 下载完整计算结果 (含明细)",
                            data=output,
                            file_name="运费计算结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("⚠️ 计算完成，但没有提取到有效的运费数据。")
                except Exception as e:
                    st.error(f"❌ 计算过程中发生错误：{str(e)}")

# ----------------------------------------
# 工具3：排料计算
# ----------------------------------------
elif menu_choice == "✂️ 排料计算 (开发中)":
    st.info("排料计算功能工程师正在熬夜开发中，敬请期待...")