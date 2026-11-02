# -*- coding: utf-8 -*-
"""
成本计算（新）—— 新输入口径的布料费 + 运费一体化计算
=====================================================

与旧版「布料费用计算」「运费计算」模块完全隔离：
  - 旧版：需要订单表 / 发货日报表 / SKU-款号映射 / 款号信息表
  - 新版：只需要「布料跟踪表」「款号档案」「FBA头程」「装箱明细」

输入（单个 Excel 工作簿，必需 4 张表）：
  布料跟踪表                      订单编号 / 产品名称 / 颜色&色号 /
                                  布料送货数量（匹）/ 金额（元）/ 下单日期
  款号档案                        款号 / 出货数（件/匹）
                                  —— 出货数为空 = 该款已淘汰，不参与计算
  FBA头程                         发货日期 / 货件编号 / 发货数量pcs / 总金额
  数据源-领星-FBA货件-装箱明细      货件单号 / MSKU / SKU / 申报量

时间窗口：统一近一年（365 天）
  基准日 = max(布料最新下单日, 头程最新发货日) 各自独立取
  注意：基准日取「原始下单日」，未送货的行（匹数为空）不影响窗口

成本口径：
  布料费用 K1 = 布料金额 ÷ 布料匹数 ÷ 成衣数 Q
     - Q 取自「款号档案 - 出货数（件/匹）」
     - 纯色 / 印花按产品名称是否含「印花」拆分（K1-s / K1-p）
  运费 = 该款分摊头程费用 ÷ 该款申报件数
     - 货件内按各 SKU 申报量占比分摊
  合计 = K1 + 运费

对外接口：
  run_cost_calculation(file) -> dict
    {
      "成本汇总":    DataFrame  (布料 21 款，含运费与合计)
      "仅有运费的款": DataFrame  (有运费但布料未算，附未算原因)
      "排除款清单":  DataFrame  (淘汰款 + 无布料款)
      "布料费用结果": DataFrame
      "运费结果":    DataFrame
      "运费分摊明细": DataFrame
      "数据诊断":    DataFrame
      "字段说明":    DataFrame
      "摘要":       dict        (款数 / 区间 / 覆盖率等，供页面展示)
    }
"""

from __future__ import annotations

import logging
from typing import Dict, Optional

import pandas as pd

from .fabric_cost_v2 import run_fabric_calculation
from .shipping_cost_v2 import run_shipping_calculation

logger = logging.getLogger(__name__)

FIELD_GLOSSARY = [
    {"字段": "款号", "含义": "服装款号，如 PPY0019"},
    {"字段": "成衣数Q1", "含义": "每匹布出多少件成衣，取「款号档案 - 出货数（件/匹）」"},
    {"字段": "布料费用K1", "含义": "布料总金额 ÷ 布料匹数 ÷ Q1，不含备布"},
    {"字段": "纯色布料费用K1-s", "含义": "产品名称不含「印花」的部分单独计算"},
    {"字段": "印花布料费用K1-p", "含义": "产品名称含「印花」的部分单独计算"},
    {"字段": "布料金额", "含义": "时间窗口内该款布料金额合计（元）"},
    {"字段": "布料匹数", "含义": "时间窗口内该款布料送货匹数合计"},
    {"字段": "货件数", "含义": "时间窗口内装载该款 SKU 的 FBA 货件数量"},
    {"字段": "发货件数", "含义": "该款在所有货件中的申报量合计（件）"},
    {"字段": "分摊运费合计", "含义": "按申报量占比分摊到该款的头程费用合计（元）"},
    {"字段": "运费(单件)", "含义": "分摊运费合计 ÷ 发货件数"},
    {"字段": "合计(布料+运费)", "含义": "布料费用K1 + 运费(单件)，不含辅料/加工费/包材"},
    {"字段": "样本提示", "含义": "发货件数低于阈值时的提示，数值仅供参考"},
    {"字段": "时间窗口", "含义": "统一近一年；布料基准日=最新下单日，运费基准日=最新发货日"},
]


def _empty() -> pd.DataFrame:
    return pd.DataFrame()


def _rewind(f):
    """Streamlit UploadedFile 多次读取前必须回到开头，路径则无需处理。"""
    try:
        f.seek(0)
    except Exception:
        pass
    return f


def run_cost_calculation(excel_file,
                         styles: Optional[list] = None,
                         q_overrides: Optional[Dict[str, float]] = None) -> dict:
    """新口径成本计算统一入口。

    :param excel_file: Excel 路径或文件对象
    :param styles: 只算这些款号；None = 全部
    :param q_overrides: {款号: 成衣数} 手动覆盖款号档案里的出货数
    """
    fab = run_fabric_calculation(
        _rewind(excel_file),
        styles=styles,
        q_overrides=q_overrides,
        use_profile_q=True,
    )
    shp = run_shipping_calculation(_rewind(excel_file), styles=styles) or {}

    df_fab = fab["布料费用结果"]
    df_ship = shp.get("运费结果", _empty())
    diag_f = fab.get("数据诊断") or {}
    diag_s = shp.get("数据诊断") or {}

    # ---------- 款号档案：判断未算原因 ----------
    registered, has_q = set(), set()
    try:
        profile = pd.read_excel(_rewind(excel_file), sheet_name="款号档案")
        profile.columns = [str(c).strip() for c in profile.columns]
        profile["款号"] = profile["款号"].astype(str).str.strip().str.upper()
        registered = set(profile["款号"])
        q_col = next((c for c in profile.columns if "出货数" in c), None)
        if q_col:
            has_q = set(profile.loc[profile[q_col].notna(), "款号"])
    except Exception as e:      # 款号档案缺失时不影响主计算
        logger.warning("读取款号档案失败，未算原因将留空: %s", e)

    # ---------- 主表：布料款 + 运费 ----------
    if df_fab.empty:
        main = _empty()
    else:
        cols = ["款号", "成衣数Q1", "布料费用K1", "纯色布料费用K1-s",
                "印花布料费用K1-p", "布料金额", "布料匹数"]
        cols = [c for c in cols if c in df_fab.columns]
        main = df_fab[cols].copy()

        ship_cols = ["款号", "货件数", "发货件数", "分摊运费合计", "运费(单件)", "样本提示"]
        ship_cols = [c for c in ship_cols if c in df_ship.columns]
        main = main.merge(df_ship[ship_cols], on="款号", how="left")

        if "运费(单件)" in main.columns:
            main["运费(单件)"] = main["运费(单件)"].round(2)
            main["合计(布料+运费)"] = (
                main["布料费用K1"] + main["运费(单件)"].fillna(0)).round(2)
            main = main.sort_values("合计(布料+运费)", ascending=False)
        main.insert(0, "序号", range(1, len(main) + 1))

    # ---------- 附表：有运费但布料未算 ----------
    if df_fab.empty or df_ship.empty:
        only_ship = _empty()
    else:
        only_ship = df_ship[~df_ship["款号"].isin(set(df_fab["款号"]))].copy()

        def reason(row):
            k = str(row["款号"]).strip().upper()
            if registered and k not in registered:
                return "SKU 档案未登记的旧命名款号"
            if has_q and k not in has_q:
                return "款号档案出货数为空（已淘汰）"
            return "有出货数但近一年无布料送货"

        if not only_ship.empty:
            only_ship["未算原因"] = only_ship.apply(reason, axis=1)
            keep = [c for c in ["款号", "货件数", "发货件数", "分摊运费合计",
                                "运费(单件)", "样本提示", "未算原因"]
                    if c in only_ship.columns]
            only_ship = only_ship[keep].sort_values(
                "分摊运费合计", ascending=False)

    # ---------- 附表：排除款清单 ----------
    excluded = diag_f.get("有布料但出货数为空(已淘汰,不计算)") or []
    no_fab = diag_f.get("有出货数但窗口内无布料数据") or []
    ex_rows = [{"款号": x, "状态": "已淘汰（款号档案出货数为空，不计算）"} for x in excluded]
    ex_rows += [{"款号": x, "状态": "有出货数，但近一年无布料送货"} for x in no_fab]
    df_ex = pd.DataFrame(ex_rows) if ex_rows else _empty()

    # ---------- 诊断 ----------
    diag_rows = [{"项目": f"[布料] {k}",
                  "值": "、".join(map(str, v)) if isinstance(v, list) else v}
                 for k, v in diag_f.items()]
    diag_rows += [{"项目": f"[运费] {k}", "值": v} for k, v in diag_s.items()]
    df_diag = pd.DataFrame(diag_rows) if diag_rows else _empty()

    # ---------- 摘要 ----------
    summary = {
        "布料款数": len(df_fab),
        "运费款数": len(df_ship),
        "完整成本款数": len(main),
        "仅有运费款数": len(only_ship),
        "排除款数": len(df_ex),
        "运费覆盖率": diag_s.get("金额覆盖率", "-"),
        "发货件数合计": int(df_ship["发货件数"].sum()) if not df_ship.empty else 0,
    }
    if not main.empty and "合计(布料+运费)" in main.columns:
        summary["合计最低"] = float(main["合计(布料+运费)"].min())
        summary["合计最高"] = float(main["合计(布料+运费)"].max())
        summary["布料区间"] = (float(main["布料费用K1"].min()),
                          float(main["布料费用K1"].max()))
        if "运费(单件)" in main.columns and main["运费(单件)"].notna().any():
            summary["运费区间"] = (float(main["运费(单件)"].min()),
                              float(main["运费(单件)"].max()))
        summary["布料窗口"] = (str(df_fab["起始日期"].iloc[0]),
                          str(df_fab["结束日期"].iloc[0]))
    if not df_ship.empty:
        summary["运费窗口"] = (str(df_ship["起始日期"].iloc[0]),
                          str(df_ship["结束日期"].iloc[0]))

    return {
        "成本汇总": main,
        "仅有运费的款": only_ship,
        "排除款清单": df_ex,
        "布料费用结果": df_fab,
        "运费结果": df_ship,
        "运费分摊明细": shp.get("df4明细数据", _empty()),
        "数据诊断": df_diag,
        "字段说明": pd.DataFrame(FIELD_GLOSSARY),
        "摘要": summary,
    }


def export_excel(result: dict, path_or_buffer):
    """把结果写成多 sheet 工作簿。path_or_buffer 可以是路径或 BytesIO。"""
    with pd.ExcelWriter(path_or_buffer, engine="openpyxl") as w:
        result["字段说明"].to_excel(w, sheet_name="字段说明", index=False)
        result["成本汇总"].to_excel(w, sheet_name="成本汇总", index=False)
        result["仅有运费的款"].to_excel(w, sheet_name="仅有运费的款", index=False)
        result["排除款清单"].to_excel(w, sheet_name="排除款清单", index=False)
        result["布料费用结果"].to_excel(w, sheet_name="布料费用结果", index=False)
        result["运费结果"].to_excel(w, sheet_name="运费结果", index=False)
        result["运费分摊明细"].to_excel(w, sheet_name="运费分摊明细", index=False)
        result["数据诊断"].to_excel(w, sheet_name="数据诊断", index=False)
    return path_or_buffer


__all__ = ["run_cost_calculation", "export_excel", "FIELD_GLOSSARY"]
