"""
运费费用计算器（v4 · 装箱明细口径）
====================================
数据输入源（单个 Excel）：
  - FBA头程：发货日期 / 货件编号 / 总金额 / 发货数量pcs
  - 数据源-领星-FBA货件-装箱明细：货件单号 / MSKU / 申报量

口径变更（2026-10-23）：
  旧：发货日报表（SKU×发货订单编号×发货数量）→ 按订单内款号数量占比分摊
  新：装箱明细（货件单号×MSKU×申报量）→ 按货件内款号申报量占比分摊
  原因：发货日报表被领星装箱明细替代，后者直接给出每个 FBA 货件里各 SKU 的件数。

合并单元格处理（重要）：
  装箱明细的「货件单号 / 装箱方式 / 总箱数 / 总重量 / 总体积 / 店铺 / 国家」
  都是以**合并单元格**的形式一个货件占多行，值只存在区域左上角。
  pandas 读出来只在首行有值，其余为 NaN → 必须向下填充(ffill)。
  已用 openpyxl 解析 1106 个合并区域验证：ffill 与真实合并展开结果 **0 行差异**。

计算逻辑：
  1. FBA头程：金额>0 + 有效发货日期
  2. 装箱明细：ffill 货件单号 → 从 MSKU 前缀提取款号 → 按(货件, 款号)汇总申报量
  3. 占货件比 = 该款申报量 / 该货件总申报量
  4. 货件金额 = 同货件所有头程行总金额之和
  5. 款运费 = Σ(货件金额 × 占货件比) / Σ该款申报量

对外接口：
  run_shipping_calculation(file, config=None, styles=None, periods=None)
    -> {"运费结果", "df4明细数据", "时间段汇总"}
"""
import logging
from datetime import datetime, timedelta
from typing import Dict, List, Optional

import pandas as pd

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# =========================================================
# 配置区：工作表名 / 列名 / 业务阈值，随模板变化只改这里
# =========================================================
CONFIG = {
    "sheets": ["FBA头程", "数据源-领星-FBA货件-装箱明细"],
    "default_lookback_days": 365,     # 默认时间窗口：近一年
    "cols": {
        "fba": {"date": "发货日期", "code": "货件编号",
                "amount": "总金额", "qty": "发货数量pcs"},
        "pack": {"code": "货件单号", "sku": "MSKU", "qty": "申报量"},
    },
    # 装箱明细中需要向下填充的合并单元格列（合并区域左上角才有值）
    "pack_ffill_cols": ["货件单号", "装箱方式", "总箱数", "总重量",
                        "总体积", "店铺", "国家"],
    "small_sample": 300,   # 件数低于此值提示样本偏小
    "tiny_sample": 100,    # 件数低于此值提示样本过小
}


def _parse_date(value, default=None):
    if value is None:
        return default
    return pd.to_datetime(value, errors="coerce")


def _norm_code(s: pd.Series) -> pd.Series:
    """货件编号标准化：去空格、转大写。"""
    return s.astype(str).str.strip().str.upper().replace(
        {"NAN": pd.NA, "NONE": pd.NA, "": pd.NA})


def _extract_style(s: pd.Series) -> pd.Series:
    """从 MSKU 提取款号：PPY0019-10-1303A → PPY0019"""
    return (s.astype(str).str.strip().str.upper()
            .str.extract(r"^\s*([A-Z]{2,3}\d+)")[0])


class ShippingCostCalculator:
    """运费费用计算器（装箱明细口径）"""

    def __init__(self, excel_file, config: dict = None,
                 styles: Optional[List[str]] = None,
                 periods: Optional[Dict[str, dict]] = None):
        self.excel_file = excel_file
        self.cfg = config or CONFIG
        self.styles = styles
        self.periods = periods or {}
        self.data = {}
        self.df1 = self.df2 = self.df3 = self.df4 = None
        self.pack_diag = {}
        self.period_summary = []

    # ---------- 数据加载与校验 ----------
    def load_data(self):
        logger.info("正在加载 Excel 文件...")
        try:
            xls = pd.ExcelFile(self.excel_file)
        except Exception as e:
            raise ValueError(f"无法读取 Excel 文件: {e}") from e

        missing = [s for s in self.cfg["sheets"] if s not in xls.sheet_names]
        if missing:
            raise ValueError(
                f"缺少工作表: {missing}。需要「FBA头程」和"
                f"「数据源-领星-FBA货件-装箱明细」两张表。")
        for sheet in self.cfg["sheets"]:
            self.data[sheet] = pd.read_excel(xls, sheet_name=sheet)
        logger.info("数据加载成功")

    @staticmethod
    def _require_cols(df: pd.DataFrame, cols: dict, sheet: str):
        missing = [c for c in cols.values() if c not in df.columns]
        if missing:
            raise ValueError(f"工作表「{sheet}」缺少列: {missing}")

    # ---------- 时间窗口 ----------
    def _default_end(self) -> pd.Timestamp:
        return getattr(self, "_data_end", None) or self.df1[
            self.cfg["cols"]["fba"]["date"]].max()

    def resolve_period(self, style: str):
        end_default = self._default_end()
        p = self.periods.get(style)
        if p:
            start = _parse_date(p.get("start"))
            end = _parse_date(p.get("end"))
            days = p.get("days")
            if start is None and end is not None and days:
                start = end - timedelta(days=days)
            if end is None and start is not None and days:
                end = start + timedelta(days=days)
            if end is None:
                end = end_default
            if start is None:
                start = end - timedelta(days=days or self.cfg["default_lookback_days"])
        else:
            end = end_default
            start = end - timedelta(days=self.cfg["default_lookback_days"])
        return start, end

    # ---------- 数据加工 ----------
    def build_df1(self):
        """FBA头程：金额>0、日期有效、货件编号标准化"""
        c = self.cfg["cols"]["fba"]
        raw = self.data["FBA头程"]
        # 窗口基准日取**全量**数据的最新发货日（不受金额过滤影响）
        _all_dates = pd.to_datetime(raw[c["date"]], errors="coerce")
        self._data_end = _all_dates.max()

        fba = raw.copy()
        fba = fba[pd.to_numeric(fba[c["amount"]], errors="coerce").fillna(0) > 0].copy()
        fba[c["date"]] = pd.to_datetime(fba[c["date"]], errors="coerce")
        fba = fba.dropna(subset=[c["date"]]).copy()
        if fba.empty:
            raise ValueError("FBA头程 中没有有效的发货日期数据")

        fba["_code"] = _norm_code(fba[c["code"]])
        fba["_amt"] = pd.to_numeric(fba[c["amount"]], errors="coerce").fillna(0)
        fba["_qty"] = pd.to_numeric(fba[c["qty"]], errors="coerce") \
            if c["qty"] in fba.columns else pd.NA
        self.df1 = fba

    def build_packing(self) -> pd.DataFrame:
        """装箱明细：处理合并单元格 → 提取款号 → 按(货件,款号)汇总申报量"""
        c = self.cfg["cols"]["pack"]
        pack = self.data["数据源-领星-FBA货件-装箱明细"].copy()

        # 合并单元格：值只在区域左上角，必须向下填充
        for col in self.cfg["pack_ffill_cols"]:
            if col in pack.columns:
                pack[col] = pack[col].ffill()

        pack["_code"] = _norm_code(pack[c["code"]])
        pack["_msku"] = pack[c["sku"]].astype(str).str.strip()
        pack["款号"] = _extract_style(pack["_msku"])
        pack["_qty"] = pd.to_numeric(pack[c["qty"]], errors="coerce")

        n_rows = len(pack)
        no_code = int(pack["_code"].isna().sum())
        no_style = int(pack["款号"].isna().sum())
        pack = pack[pack["_code"].notna() & pack["款号"].notna()
                    & (pack["_qty"].fillna(0) > 0)].copy()

        # 按 (货件, 款号) 汇总
        g = (pack.groupby(["_code", "款号"], as_index=False)
             .agg(申报量=("_qty", "sum"),
                  SKU数=("_msku", "nunique"),
                  明细行数=("_qty", "size")))
        g["货件总申报量"] = g.groupby("_code")["申报量"].transform("sum")
        g["占货件比"] = g["申报量"] / g["货件总申报量"].replace(0, pd.NA)

        self.pack_diag = {
            "装箱明细总行数": n_rows,
            "货件单号缺失行": no_code,
            "款号无法识别行": no_style,
            "有效货件数": int(pack["_code"].nunique()),
            "有效明细行": len(pack),
            "涉及款号数": int(pack["款号"].nunique()),
        }
        return g

    def build_df2(self):
        """货件金额 ∪ 装箱占比：逐(货件,款号)分摊金额"""
        g = self.build_packing()

        # 货件 → 金额 / 日期（同货件多头程行：金额求和，日期取最早）
        c = self.cfg["cols"]["fba"]
        h = (self.df1.groupby("_code", as_index=False)
             .agg(货件金额=("_amt", "sum"),
                  发货日期=(c["date"], "min"),
                  头程件数=("_qty", lambda s: pd.to_numeric(s, errors="coerce").sum())))
        self.df_h = h

        m = g.merge(h, on="_code", how="inner")
        m["分摊金额"] = m["货件金额"] * m["占货件比"]

        # 货件完整度诊断：装箱申报量 / 头程发货量
        m["完整度"] = (m["货件总申报量"] / m["头程件数"].replace(0, pd.NA)).round(2)
        self.pack_diag["可匹配头程的货件"] = int(m["_code"].nunique())
        self.pack_diag["头程有但无装箱明细的货件"] = int(
            len(set(self.df1["_code"].dropna()) - set(g["_code"].dropna())))
        self.df2 = m

    def build_df4(self, start, end):
        """按指定窗口过滤"""
        return self.df2[(self.df2["发货日期"] >= start) &
                        (self.df2["发货日期"] <= end)].copy()

    def _target_styles(self) -> List[str]:
        available = set(self.df2["款号"].dropna().unique())
        if self.styles:
            return [s for s in self.styles if s in available]
        return sorted(available)

    def calculate_final(self):
        rows = []
        for style in self._target_styles():
            start, end = self.resolve_period(style)
            d = self.build_df4(start, end)
            d = d[d["款号"] == style]
            if d.empty:
                continue
            qty = d["申报量"].sum()
            if qty <= 0:
                continue
            amt = d["分摊金额"].sum()
            note = ""
            if qty < self.cfg["tiny_sample"]:
                note = "样本过小，仅供参考"
            elif qty < self.cfg["small_sample"]:
                note = "样本偏小"
            rows.append({
                "款号": style,
                "起始日期": start.strftime("%Y-%m-%d"),
                "结束日期": end.strftime("%Y-%m-%d"),
                "货件数": int(d["_code"].nunique()),
                "发货件数": int(qty),
                "分摊运费合计": round(amt, 2),
                "运费(单件)": round(amt / qty, 2),
                "样本提示": note,
            })
            self.period_summary.append({"款号": style,
                                        "起始日期": start.strftime("%Y-%m-%d"),
                                        "结束日期": end.strftime("%Y-%m-%d")})
        if not rows:
            return pd.DataFrame(columns=["款号", "起始日期", "结束日期",
                                         "货件数", "发货件数",
                                         "分摊运费合计", "运费(单件)", "样本提示"])
        return pd.DataFrame(rows).sort_values("分摊运费合计", ascending=False) \
            .reset_index(drop=True)

    def coverage(self):
        """数据覆盖诊断：窗口内有多少头程金额能配上装箱明细"""
        c = self.cfg["cols"]["fba"]
        end = self._default_end()
        start = end - timedelta(days=self.cfg["default_lookback_days"])
        fh_w = self.df1[(self.df1[c["date"]] >= start) &
                        (self.df1[c["date"]] <= end)]
        codes = set(self.df2["_code"].dropna()) if self.df2 is not None else set()
        hit = fh_w["_code"].isin(codes)
        return {
            "窗口起": start.strftime("%Y-%m-%d"),
            "窗口止": end.strftime("%Y-%m-%d"),
            "窗口内头程行": int(len(fh_w)),
            "窗口内头程金额": round(float(fh_w["_amt"].sum()), 2),
            "有装箱明细的头程行": int(hit.sum()),
            "有装箱明细的金额": round(float(fh_w.loc[hit, "_amt"].sum()), 2),
            "金额覆盖率": f"{(fh_w.loc[hit, '_amt'].sum() / fh_w['_amt'].sum() * 100 if fh_w['_amt'].sum() else 0):.1f}%",
        }

    def calculate(self):
        self.load_data()
        c = self.cfg["cols"]
        self._require_cols(self.data["FBA头程"], c["fba"], "FBA头程")
        self._require_cols(self.data["数据源-领星-FBA货件-装箱明细"],
                           c["pack"], "数据源-领星-FBA货件-装箱明细")
        self.build_df1()
        self.build_df2()
        return self.calculate_final()


def run_shipping_calculation(file, config: dict = None,
                             styles: Optional[List[str]] = None,
                             periods: Optional[Dict[str, dict]] = None):
    """统一入口。
    返回：{"运费结果", "df4明细数据", "时间段汇总", "数据诊断"}
    """
    calculator = ShippingCostCalculator(file, config, styles=styles, periods=periods)
    final_results = calculator.calculate()
    if final_results.empty:
        return None
    return {
        "运费结果": final_results,
        "df4明细数据": calculator.df2,
        "时间段汇总": pd.DataFrame(calculator.period_summary),
        "数据诊断": {**calculator.pack_diag, **calculator.coverage()},
    }


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python shipping_cost.py <Excel路径> [款号1,款号2...]")
        sys.exit(1)
    styles = sys.argv[2].split(",") if len(sys.argv) > 2 else None
    out = run_shipping_calculation(sys.argv[1], styles=styles)
    if out:
        print(out["运费结果"].to_string(index=False))
        print("\n[数据诊断]")
        for k, v in out["数据诊断"].items():
            print(f"  {k}: {v}")
    else:
        print("无有效计算结果")
