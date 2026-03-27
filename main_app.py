import streamlit as st
import pandas as pd
import io
# 这里就是调用你的后厨脚本
from fabric_cost import run_fabric_calculation

# 设置网页全屏宽一点，好看一点
st.set_page_config(page_title="公司数据处理平台", layout="wide")

st.title("📊 自动化计算工具平台")

# 创建左侧边栏，作为菜单
st.sidebar.title("工具菜单")
# 这里设计了下拉菜单，以后写了新脚本，直接在这里加名字
menu_choice = st.sidebar.radio(
    "请选择你需要使用的功能：",
    ["🧵 布料费用计算", "🚚 运费计算 (开发中)", "✂️ 排料计算 (开发中)"]
)

# ----------------------------------------
# 工具1：布料费用计算
# ----------------------------------------
if menu_choice == "🧵 布料费用计算":
    st.header("布料费用自动化计算")
    st.write("请上传包含【布料跟踪表, 订单表, SKU-款号映射, 款号信息表】的 Excel 文件")

    # 1. 提供上传文件的按钮
    uploaded_file = st.file_uploader("点击此处上传 Excel", type=["xlsx", "xls"])

    # 2. 如果传了文件，显示计算按钮
    if uploaded_file is not None:
        if st.button("🚀 开始计算"):
            # 出现加载动画
            with st.spinner("程序正在疯狂计算中，请耐心等待几秒钟..."):
                try:
                    # 调用 fabric_cost.py 里的函数
                    result_df = run_fabric_calculation(uploaded_file)

                    if result_df is not None and not result_df.empty:
                        st.success("🎉 计算成功！")
                        # 在网页上直接显示表格
                        st.dataframe(result_df)

                        # 把表格转换为可以下载的 Excel 文件流
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_df.to_excel(writer, index=False)
                        output.seek(0)

                        # 提供一键下载按钮
                        st.download_button(
                            label="📥 下载计算结果 (Excel格式)",
                            data=output,
                            file_name="布料费用计算结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("⚠️ 计算完成，但没有提取到符合条件的有效数据。")
                except Exception as e:
                    st.error(f"❌ 计算过程中发生错误，请检查表格格式：{str(e)}")

# ----------------------------------------
# 工具2：未来可以增加的运费计算
# ----------------------------------------
elif menu_choice == "🚚 运费计算 (开发中)":
    st.info("运费计算功能工程师正在熬夜开发中，敬请期待...")