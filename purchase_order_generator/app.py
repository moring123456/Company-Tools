# -*- coding: utf-8 -*-
"""app.py — 采购单自动生成系统 独立入口

运行: streamlit run purchase_order_generator/app.py
（也可被 Company-Tools 的 main_app.py 单入口集成，见 ui.py）
"""
import os
import sys

# 保证模块可被绝对导入（平台集成 / 独立运行均可）
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st

from purchase_order_generator.ui import render_purchase_order_page

st.set_page_config(page_title="采购单自动生成系统", page_icon="📦", layout="wide")
st.title("📦 采购单自动生成系统")

render_purchase_order_page()
