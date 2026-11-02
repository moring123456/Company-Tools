"""
布料费用计算器（v7 · 成衣数手填口径）
======================================
支持两种输入模式：

【模式 A · profile_q（推荐，2026-10-23 起启用）】
  数据输入源（2 个工作表）：
    - 布料跟踪表: 订单编号 / 产品名称 / 颜色&色号 / 布料送货数量（匹） / 金额（元） / 下单日期
    - 款号档案:   款号 / 出货数（件/匹）   ← 成衣数 Q 由业务手工填写
  特点：
    - 不再需要「订单表」「收货日报表」——Q 直接取款号档案
    - 款号档案中「出货数」为空的款 = **已淘汰**，直接跳过，不参与计算（不是数据缺失）
    - K1 = Σ布料金额 ÷ Σ送货匹数 ÷ Q（纯色/印花分桶算 K1-s / K1-p）
    - 混搭明细不可用（混搭识别依赖订单表颜色列）

【模式 B · 完整订单表模式（历史版本，保留兼容）】
  数据输入源（4 个工作表）：布料跟踪表 / 订单表 / SKU-款号映射 / 款号信息表
  Q1 = 已收货数量 ÷ 布料送货匹数；支持备布折算与混搭 K1-m

时间窗口：
  - 默认：最近 365 天（CONFIG.default_lookback_days 可调）
  - 自定义：periods = {"PPY0019": {"start": "2026-01-01", "end": "2026-06-30"}}

核心公式：
  K1（单件布料成本）= 布料总金额 C1 / 布料匹数 N1 / 成衣数 Q
  纯色 K1-s、印花 K1-p 分别按产品名称是否含「印花」分桶计算

对外接口：
  run_fabric_calculation(file, config=None, styles=None, periods=None,
                         q_overrides=None, treat_mixed_as_print=False,
                         use_profile_q=False)
    -> {"布料费用结果": df, "备布数据": df, "时间段汇总": df, "混搭明细": df, "数据诊断": dict}
"""
import logging
import re
from datetime import datetime, timedelta
from typing import Dict, List, Optional

import pandas as pd

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# =========================================================
# 配置区：工作表名 / 列名 / 业务阈值，随模板变化只改这里
# =========================================================
CONFIG = {
    "sheets": ["布料跟踪表", "订单表", "SKU-款号映射", "款号信息表"],
    # 模式 A：只用这两张表，Q 取款号档案的「出货数（件/匹）」
    "profile_q_sheets": ["布料跟踪表", "款号档案"],
    "profile_q_cols": {"style": "款号", "q": "出货数（件/匹）"},
    "receipt_sheet": "收货日报表",   # 可选：用收货日报表重算已收货数量（订单表已收货常滞后为0）
    "default_lookback_days": 365,    # 默认时间窗口：近一年（v4 从 180 改为 365）
    "receipt_ratio": 0.70,           # 已收货数量 >= 下单数量 × 70% 才视为有效
    "treat_unreceived_as_backup": True,  # 未收货（已收货=0）的布料行按备布折算处理，而非剔除
    "sku_style_prefix": r"^(([A-Z]{2,3})\d+)",    # 从 SKU/产品名 提取款号的正则：2~3个大写字母 + 1位以上数字（如 PPY0075 / BTE1001 / ADM42）
    "color_prefixes": ["40STR汗布", "40STR", "32STR汗布", "32STR"],
    "mix_sub_color_indicators": ["YH\\d+\\.?\\d*", "G\\d+-?\\d+"],  # 印花子色的标识符
    "mix_separator": "+",             # 混搭内部子色分隔符
    "cols": {
        "fabric": {
            "order": "订单编号", "product": "产品名称", "color": "颜色&色号",
            "rolls": "布料送货数量（匹）", "amount": "金额（元）", "date": "下单日期",
        },
        "order": {
            "order": "订单编号", "color": "颜色", "sku": "SKU编码",
            "qty": "下单数量", "received": "已收货数量",
        },
        "receipt": {
            "order": "订单编号", "sku": "SKU编码", "qty": "收货数量",
        },
        "mapping": {"sku": "SKU", "style": "款号"},
        "style_info": {"style": "款号", "position": "产品定位"},
    },
    "eliminated_position": "淘汰款",
}


def _parse_date(value, default=None):
    if value is None:
        return default
    return pd.to_datetime(value, errors="coerce")


class FabricCostCalculator:
    """布料费用计算器（数据输入：Excel 文件对象或路径）"""

    def __init__(self, excel_file, config: dict = None,
                 styles: Optional[List[str]] = None,
                 periods: Optional[Dict[str, dict]] = None,
                 q_overrides: Optional[Dict[str, float]] = None,
                 treat_mixed_as_print: bool = False,
                 use_profile_q: bool = False):
        """
        :param excel_file: Excel 路径或文件对象
        :param config: 覆盖默认 CONFIG
        :param styles: 只计算这些款号；None=全部非淘汰款
        :param periods: {款号: {"start":..,"end":..}} 自定义时间段
        :param q_overrides: {款号: 成衣数} 人工核定成衣数（优先级最高）
        :param treat_mixed_as_print: True=混搭并入印花体系（仅模式 B 有效）
        :param use_profile_q: True=模式 A，Q 取款号档案「出货数（件/匹）」，无需订单表
        """
        self.excel_file = excel_file
        self.cfg = config or CONFIG
        self.styles = styles
        self.periods = periods or {}
        self.q_overrides = q_overrides or {}
        self.treat_mixed_as_print = treat_mixed_as_print
        self.use_profile_q = use_profile_q
        self.data = {}
        self.period_summary = []
        self.quality_report = []
        self.mixed_results_all = []
        self.diag = {}

    # ---------- 数据加载与校验 ----------
    def load_data(self):
        logger.info("正在读取文件并加载工作表...")
        try:
            xls = pd.ExcelFile(self.excel_file)
        except Exception as e:
            raise ValueError(f"无法读取 Excel 文件: {e}") from e

        sheets = self.cfg["profile_q_sheets"] if self.use_profile_q else self.cfg["sheets"]
        missing = [s for s in sheets if s not in xls.sheet_names]
        if missing:
            if self.use_profile_q:
                raise ValueError(
                    f"缺少工作表: {missing}。手填成衣数模式需要「布料跟踪表」"
                    f"和「款号档案（含『出货数（件/匹）』列）」两张表。")
            raise ValueError(f"缺少工作表: {missing}，请确认包含 {self.cfg['sheets']}")
        for sheet in sheets:
            self.data[sheet] = pd.read_excel(xls, sheet_name=sheet)

        if self.use_profile_q:
            logger.info("数据加载成功（手填成衣数模式，跳过订单表）")
            return
        # 数据质量检查：关键表去重（必须在收货日报表重算之前执行）
        self._data_quality_check()
        logger.info("数据加载成功")
        self._sync_received_from_receipt_sheet()

    def _data_quality_check(self):
        """【必要检查流程】关键表去重与重复行体检。

        背景：实测源数据「订单表」存在 26 个订单整单重复（每行 ×2 完全重复，
        共 1019 组 (订单,SKU)）。不去重会把收货/下单数量放大一倍，导致成衣数 Q
        严重虚高（PPY0075 自动 Q 从 114.18 修正为 63.66）。本流程对所有关键表
        做去重并生成质量报告，供用户核验源数据。

        检查规则：
          订单表        按 (订单编号, SKU编码) 去重（业务上订单+SKU 应唯一）
          布料跟踪表    全列完全重复行去重
          SKU-款号映射  按 SKU 去重
        """
        self.quality_report = []
        oc = self.cfg["cols"]["order"]

        # 1) 订单表：按 (订单编号, SKU编码) 去重
        order_df = self.data["订单表"]
        n0 = len(order_df)
        order_df = order_df.drop_duplicates(subset=[oc["order"], oc["sku"]], keep="first")
        self.data["订单表"] = order_df
        self._log_quality("订单表", f"({oc['order']}, {oc['sku']})", n0, len(order_df))

        # 2) 布料跟踪表：全列完全重复行去重
        if "布料跟踪表" in self.data:
            fab = self.data["布料跟踪表"]
            n0 = len(fab)
            fab = fab.drop_duplicates(keep="first")
            self.data["布料跟踪表"] = fab
            self._log_quality("布料跟踪表", "全列", n0, len(fab))

        # 3) SKU-款号映射：按 SKU 去重
        if "SKU-款号映射" in self.data:
            mc = self.cfg["cols"]["mapping"]
            m = self.data["SKU-款号映射"]
            n0 = len(m)
            m = m.drop_duplicates(subset=[mc["sku"]], keep="first")
            self.data["SKU-款号映射"] = m
            self._log_quality("SKU-款号映射", mc["sku"], n0, len(m))

    def _log_quality(self, sheet, key, n_before, n_after):
        n_del = n_before - n_after
        rec = {"工作表": sheet, "检查键": key,
               "原始行数": n_before, "去重后行数": n_after, "删除重复行": n_del}
        self.quality_report.append(rec)
        if n_del > 0:
            logger.warning(
                f"数据质量检查「{sheet}」：{n_before} → {n_after} 行，"
                f"按 {key} 删除 {n_del} 行重复（请核验源数据）")

    def _sync_received_from_receipt_sheet(self):
        """用收货日报表重算订单表的已收货数量。

        背景：订单表的「已收货数量」字段经常滞后（很多行是 0），但收货日报表
        里有逐笔实际收货记录。两者按 (订单编号, SKU编码) 关联，把收货日报表
        的收货数量汇总后覆盖订单表的已收货数量，解决已收货=0 导致成本失真。
        """
        receipt_sheet = self.cfg.get("receipt_sheet")
        if not receipt_sheet:
            return
        xls = pd.ExcelFile(self.excel_file)
        if receipt_sheet not in xls.sheet_names:
            logger.warning(f"未找到工作表「{receipt_sheet}」，跳过已收货重算")
            return
        rc = self.cfg["cols"]["receipt"]
        oc = self.cfg["cols"]["order"]
        receipt = pd.read_excel(xls, sheet_name=receipt_sheet)
        missing = [c for c in rc.values() if c not in receipt.columns]
        if missing:
            logger.warning(f"收货日报表缺少列 {missing}，跳过已收货重算")
            return

        receipt = receipt.copy()
        receipt[rc["qty"]] = pd.to_numeric(receipt[rc["qty"]], errors="coerce").fillna(0)
        recv_agg = receipt.groupby([rc["order"], rc["sku"]], dropna=False)[rc["qty"]].sum().reset_index()

        order = self.data["订单表"].copy()
        order[oc["order"]] = order[oc["order"]].astype(str).str.strip()
        order[oc["sku"]] = order[oc["sku"]].astype(str).str.strip()
        recv_agg[rc["order"]] = recv_agg[rc["order"]].astype(str).str.strip()
        recv_agg[rc["sku"]] = recv_agg[rc["sku"]].astype(str).str.strip()

        recv_map = recv_agg.rename(columns={
            rc["order"]: "_r_order", rc["sku"]: "_r_sku", rc["qty"]: "_r_qty"})
        merged = order.merge(recv_map, left_on=[oc["order"], oc["sku"]],
                             right_on=["_r_order", "_r_sku"], how="left")
        merged[oc["received"]] = merged["_r_qty"].fillna(merged[oc["received"]])
        merged = merged.drop(columns=["_r_order", "_r_sku", "_r_qty"])
        self.data["订单表"] = merged
        logger.info(f"已用收货日报表重算已收货数量（覆盖 {len(recv_agg)} 组收货记录）")

    @staticmethod
    def _require_cols(df: pd.DataFrame, cols: dict, sheet: str):
        missing = [c for c in cols.values() if c not in df.columns]
        if missing:
            raise ValueError(f"工作表「{sheet}」缺少列: {missing}")

    # ---------- 时间窗口解析 ----------
    def _style_dates(self, style_no: str) -> pd.DataFrame:
        """该款全部布料记录（含备布）的 (下单日期, 订单编号) 去重表"""
        c = self.cfg["cols"]["fabric"]
        fabric = self.data["布料跟踪表"].copy()
        fabric[c["date"]] = pd.to_datetime(fabric[c["date"]], errors="coerce")
        fabric = fabric.dropna(subset=[c["date"]])
        if fabric.empty:
            return pd.DataFrame()
        fabric["_style_by_product"] = fabric[c["product"]].map(self._extract_style_from_product_name)
        style_rows = fabric[fabric["_style_by_product"] == style_no]
        return style_rows[[c["date"], c["order"]]].drop_duplicates()

    def _default_end(self, style_no: str) -> pd.Timestamp:
        """默认窗口结束日 = 该款最近一次布料下单日期（无则用数据全局最近日期）"""
        dates = self._style_dates(style_no)
        if not dates.empty:
            return dates[self.cfg["cols"]["fabric"]["date"]].max()
        fabric = self.data["布料跟踪表"].copy()
        fabric[self.cfg["cols"]["fabric"]["date"]] = pd.to_datetime(
            fabric[self.cfg["cols"]["fabric"]["date"]], errors="coerce")
        return fabric[self.cfg["cols"]["fabric"]["date"]].max()

    def resolve_period(self, style_no: str):
        """返回该款实际使用的 (start, end)。规则与运费计算器一致。"""
        end_default = self._default_end(style_no)
        p = self.periods.get(style_no)
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

    # ---------- 颜色 / 款号 解析（向量化辅助） ----------
    def _strip_prefix(self, s: str) -> str:
        s = str(s).strip()
        for p in self.cfg["color_prefixes"]:
            if s.startswith(p):
                s = s[len(p):].strip()
                break
        return s

    def extract_color_name(self, color_str) -> Optional[str]:
        if pd.isna(color_str):
            return None
        s = self._strip_prefix(color_str)
        m = re.sub(r"^(?:#?\d+\s*#?\s*|\d+\s*#\s*)", "", s)
        return m.strip() or (s.strip() or None)

    def extract_color_number(self, color_str) -> Optional[str]:
        if pd.isna(color_str):
            return None
        s = self._strip_prefix(color_str)
        m = re.search(r"#(\w+)", s)
        if m:
            return "#" + m.group(1)
        m2 = re.search(r"^(\d+)\s*#", s)
        return "#" + m2.group(1) if m2 else None

    def extract_print_code(self, color_str) -> Optional[str]:
        if pd.isna(color_str):
            return None
        m = re.search(r"(YH[\d.]+)", str(color_str))
        return m.group(1) if m else None

    @staticmethod
    def _is_print(product_name) -> bool:
        if pd.isna(product_name):
            return False
        p = str(product_name).upper()
        return "印" in p and "印字" not in p

    @staticmethod
    def _is_solid(product_name) -> bool:
        return not FabricCostCalculator._is_print(product_name)

    def _extract_style_from_sku(self, sku) -> Optional[str]:
        if pd.isna(sku):
            return None
        m = re.match(self.cfg["sku_style_prefix"], str(sku).strip().upper())
        return m.group(1) if m else None

    def _extract_style_from_product_name(self, product_name) -> Optional[str]:
        if pd.isna(product_name):
            return None
        m = re.match(self.cfg["sku_style_prefix"], str(product_name).strip().upper())
        return m.group(1) if m else None

    # ---------- 混搭颜色解析（v4 新增）----------
    @staticmethod
    def _is_sub_color_print(sub_color: str) -> bool:
        r"""判断子色是否是印花 \u2014\u2014 看是否含 YH\d+ / G\d+-?\d+ 等印花标识"""
        for pattern in ["YH\\d+\\.?\\d*", "G\\d+-?\\d*"]:
            if re.search(pattern, sub_color, re.IGNORECASE):
                return True
        return False

    def _parse_mix_colors(self, color_str) -> Optional[dict]:
        """拆分并识别混搭颜色字符串，返回 {sub_colors: [(name, is_print), ...]} 或 None。

        规则：
          - 用 "+" 分隔多个子色
          - 每个子色通过正则检测印花标识（YHxxx / Gxxx）判定纯色素/印花
        """
        if pd.isna(color_str):
            return None
        parts_raw = [p.strip() for p in str(color_str).split(self.cfg["mix_separator"]) if p.strip()]
        if len(parts_raw) < 2:
            return None  # 不是混搭

        sub_colors = []
        for part in parts_raw:
            name = self.extract_color_name(part) or part
            is_print = self._is_sub_color_print(part)
            sub_colors.append((name, is_print))
        return {"sub_colors": sub_colors, "n_parts": len(sub_colors)}

    def _get_mixed_ratios(self, df_input: pd.DataFrame, style_no: str) -> Dict[str, List[float]]:
        """获取该款在 DataFrame 中所有混搭行的比例权重。

        来源优先级：
        1) 尝试从输入模板的「混搭比例」列读取
        2) 兜底：等比分配

        返回：{订单编号（去点后）: [w1, w2, ...]}
        """
        c = self.cfg["cols"]["fabric"]
        result = {}

        if c["color"] not in df_input.columns:
            return result

        for _, row in df_input.iterrows():
            mix_info = self._parse_mix_colors(row[c["color"]])
            if not mix_info:
                continue  # 非混搭，跳过

            ord_num = str(row[c["order"]]).replace(".", "")
            n = mix_info["n_parts"]

            # 尝试从「混搭比例」列读取
            ratio_col = "混搭比例"
            raw_ratio = row.get(ratio_col)
            ratios = None
            if pd.notna(raw_ratio) and str(raw_ratio).strip():
                try:
                    ratios = [float(x) / sum(float(x) for x in str(raw_ratio).split(":"))
                              for x in str(raw_ratio).split(":")]
                except ValueError:
                    pass

            # 兜底：等比
            if ratios is None:
                ratios = [1.0 / n] * n

            result[ord_num] = ratios

        return result

    # ---------- 备布数据 ----------
    def build_df0_prepare(self, style_no: str, start, end) -> pd.DataFrame:
        """备布数据：无对应订单、金额>0、且下单日期落在该款时间窗口内的布料"""
        c = self.cfg["cols"]
        order_df = self.data["订单表"]
        valid_orders = set(order_df[c["order"]["order"]].dropna().astype(str).str.strip())
        fabric = self.data["布料跟踪表"].copy()
        fabric[c["fabric"]["amount"]] = pd.to_numeric(fabric[c["fabric"]["amount"]], errors="coerce")
        fabric[c["fabric"]["date"]] = pd.to_datetime(fabric[c["fabric"]["date"]], errors="coerce")
        fabric = fabric[fabric[c["fabric"]["product"]].map(
            lambda p: self._extract_style_from_product_name(p) == style_no)].copy()
        order_nums = fabric[c["fabric"]["order"]].astype(str).str.strip()
        df0 = fabric[
            (~order_nums.isin(valid_orders)) &
            (fabric[c["fabric"]["amount"]] > 0) &
            (fabric[c["fabric"]["date"]] >= start) &
            (fabric[c["fabric"]["date"]] <= end)
        ].copy()
        df0["备布标记"] = "备布"
        return df0

    # ---------- 主匹配：向量化 merge，替代原 iterrows ----------
    def build_df1(self) -> pd.DataFrame:
        c = self.cfg["cols"]
        fc = c["fabric"]       # 面料列映射（v4 修复：避免用 c["color"] 直接访问）
        oc = c["order"]        # 订单列映射
        mc = c["mapping"]      # 映射列映射
        fabric = self.data["布料跟踪表"].copy()
        order = self.data["订单表"].copy()
        mapping = self.data["SKU-款号映射"]

        # 1) 订单表预处理：清洗订单号（去点号）+ 解析颜色三要素
        order["_ord"] = order[oc["order"]].astype(str).str.replace(".", "", regex=False)
        order["_name"] = order[oc["color"]].map(self.extract_color_name)
        order["_num"] = order[oc["color"]].map(self.extract_color_number)
        order["_print"] = order[oc["color"]].map(self.extract_print_code)
        order[oc["qty"]] = pd.to_numeric(order[oc["qty"]], errors="coerce").fillna(0)
        order[oc["received"]] = pd.to_numeric(order[oc["received"]], errors="coerce").fillna(0)

        # 2) 布料表预处理
        fabric["_ord"] = fabric[fc["order"]].astype(str).str.replace(".", "", regex=False)
        fabric["_is_print"] = fabric[fc["product"]].map(self._is_print)
        fabric["_name"] = fabric[fc["color"]].map(self.extract_color_name)
        fabric["_num"] = fabric[fc["color"]].map(self.extract_color_number)
        fabric["_print"] = fabric[fc["color"]].map(self.extract_print_code)
        fabric["_orig"] = fabric[fc["color"]].map(
            lambda x: self._strip_prefix(x) if not pd.isna(x) else None)
        fabric["_row"] = range(len(fabric))
        fabric["_fabric_style"] = fabric[fc["product"]].astype(str).str.extract(
            self.cfg["sku_style_prefix"])[0]
        order["_order_style"] = order[oc["sku"]].astype(str).str.extract(
            self.cfg["sku_style_prefix"])[0]

        # ---------- 新增：混搭色解析标记（v5 改进：基于订单表颜色"+"检测混搭）----------
        # 核心逻辑：订单表颜色名含"+"就是混搭（如"黑色+黑蕾丝"/"#3 黑色+G01-6 绿格子"）
        ord_has_plus = order[oc["color"]].astype(str).str.contains(r"\+", na=False)
        mix_order_nums = set(
            order[ord_has_plus][oc["order"]].astype(str).str.replace(".", "", regex=False).str.strip().unique()
        )

        # 面料表中按订单号标记是否属于混配色订单
        def _is_in_mix_order(ord_val):
            return str(ord_val).replace(".", "").strip() in mix_order_nums

        fabric["_in_mix_order"] = fabric[fc["order"]].apply(_is_in_mix_order)
        n_mix_orders = len(mix_order_nums)
        logger.info(f"[MIX] Order-level detection: {n_mix_orders} orders contain '+' in color names")

        # 对于属于混配色订单的面料行，需要解析子色名称
        # 策略：从该订单所有行的订单颜色中提取子色列表（按"+"拆分），然后映射到面料行
        if n_mix_orders > 0:
            # 构建 {订单号: [子色名列表]} 的映射表
            mix_color_map = {}
            for _, row in order[ord_has_plus].iterrows():
                ord_clean = str(row[oc["order"]]).replace(".", "").strip()
                color_str = str(row[oc["color"]])
                parts = [p.strip() for p in color_str.split("+") if p.strip()]
                if len(parts) >= 2:
                    if ord_clean not in mix_color_map:
                        mix_color_map[ord_clean] = []
                    # 合并（去重但保留顺序）
                    existing = [e.lower() for e in mix_color_map[ord_clean]]
                    for p in parts:
                        if p.lower() not in existing:
                            mix_color_map[ord_clean].append(p)
            
            # 为面料行设置 _mix_parsed 和 _mix_sub_colors
            fabric["_mix_parsed"] = False
            fabric["_mix_sub_colors"] = None
            
            for idx in fabric[fabric["_in_mix_order"]].index:
                ord_clean = str(fabric.loc[idx, fc["order"]]).replace(".", "").strip()
                if ord_clean in mix_color_map:
                    sub_color_names = mix_color_map[ord_clean]
                    # 判断每个子色是否是印花（含 YHxxx / Gxxx）
                    sub_colors = []
                    for sc in sub_color_names:
                        is_print = self._is_sub_color_print(sc)
                        clean_name = self.extract_color_name(sc) or sc
                        sub_colors.append((clean_name, is_print))
                    
                    fabric.at[idx, "_mix_parsed"] = True
                    fabric.at[idx, "_mix_sub_colors"] = sub_colors
            
            n_mix_rows = (fabric["_mix_parsed"]).sum()
            logger.info(f"[MIX] Fabric rows marked as mixed: {n_mix_rows}")
        else:
            fabric["_mix_parsed"] = False
            fabric["_mix_sub_colors"] = None

        # 3) 匹配优先级：纯色先按颜色名→再按颜色编号；印花先按印花色号→再按颜色名
        order_keep = order[[
            "_ord", "_name", "_num", "_print", "_order_style",
            oc["sku"], oc["qty"], oc["received"],
        ]].copy()

        def merge_then_recover(df, left_key, right_key):
            right = order_keep.rename(columns={
                oc["sku"]: "_sku_r",
                oc["qty"]: "_qty_r",
                oc["received"]: "_recv_r",
            })
            # 排除 NaN 匹配键
            df_valid = df[df[left_key].notna()]
            if df_valid.empty:
                return df.iloc[0:0], df
            m = df_valid[["_row", "_ord", left_key, "_fabric_style"]].merge(
                right, left_on=["_ord", left_key], right_on=["_ord", right_key],
                how="left")
            hit = m["_sku_r"].notna() & (m["_fabric_style"] == m["_order_style"])
            return m[hit], df[~df["_row"].isin(m.loc[hit, "_row"])]

        solid = fabric[~fabric["_is_print"]]
        print_df = fabric[fabric["_is_print"]]

        parts = []
        m_solid, rest_solid = merge_then_recover(solid, "_name", "_name")
        parts.append(m_solid)
        if not rest_solid.empty:
            m_solid_fb, _ = merge_then_recover(rest_solid, "_num", "_num")
            parts.append(m_solid_fb)
        m_print, rest_print = merge_then_recover(print_df, "_print", "_print")
        parts.append(m_print)
        if not rest_print.empty:
            m_print_fb, _ = merge_then_recover(rest_print, "_name", "_name")
            parts.append(m_print_fb)

        matched = pd.concat([p for p in parts if not p.empty], ignore_index=True)

        # 4) 聚合
        agg = matched.groupby("_row").agg(
            _qty=("_qty_r", "sum"),
            _received=("_recv_r", "sum"),
            _sku=("_sku_r", "first"),
        ).reset_index()

        fabric = fabric.merge(agg, on="_row", how="left")
        fabric["下单数量"] = fabric["_qty"].fillna(0)
        fabric["已收货数量"] = fabric["_received"].fillna(0)
        fabric["颜色匹配状态"] = fabric["_sku"].where(fabric["_sku"].notna(), "匹配失败")
        fabric["颜色匹配状态"] = fabric["颜色匹配状态"].map(
            lambda x: "匹配成功" if x != "匹配失败" else "匹配失败")

        # 5) 款号
        fabric["款号"] = fabric["_sku"].map(self._extract_style_from_sku)
        mapping = mapping.drop_duplicates(subset=[c["mapping"]["sku"]])
        map_dict = dict(zip(mapping[c["mapping"]["sku"]], mapping[c["mapping"]["style"]]))
        need_map = fabric["款号"].isna() & fabric["_sku"].notna()
        fabric.loc[need_map, "款号"] = fabric.loc[need_map, "_sku"].map(map_dict)

        return fabric[fabric["颜色匹配状态"] == "匹配成功"].copy()

    def build_df1_1(self, df1: pd.DataFrame) -> pd.DataFrame:
        c = self.cfg["cols"]["fabric"]
        df1 = df1.copy()
        df1[c["amount"]] = pd.to_numeric(df1[c["amount"]], errors="coerce")
        return df1[df1[c["amount"]].notna() & (df1[c["amount"]] > 0)]

    def build_df1_2(self, df1_1: pd.DataFrame) -> pd.DataFrame:
        c = self.cfg["cols"]
        df = df1_1.copy()
        for col in [c["order"]["received"], c["order"]["qty"], c["fabric"]["rolls"]]:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        if self.cfg.get("treat_unreceived_as_backup", False):
            df["未收货标记"] = ~(df[c["order"]["received"]].fillna(0) > 0)
            has_recv = df[c["order"]["received"]].fillna(0) > 0
            df = df[has_recv & (df[c["order"]["received"]] >= df[c["order"]["qty"]].fillna(0) * self.cfg["receipt_ratio"]) | ~has_recv]
        else:
            df["未收货标记"] = False
            df = df[df[c["order"]["received"]].notna() & (df[c["order"]["received"]] > 0)]
            df = df[df[c["order"]["received"]] >= df[c["order"]["qty"]].fillna(0) * self.cfg["receipt_ratio"]]
        return df[df["款号"].notna()]

    # ---------- 款号维度计算 ----------
    def build_df2(self, df1_2: pd.DataFrame, style_no: str, start, end) -> pd.DataFrame:
        """取该款落在 [start, end] 时间窗内的布料订单，按(产品,颜色,订单)聚合"""
        cfg = self.cfg["cols"]
        c = cfg["fabric"]
        style_df = df1_2[df1_2["款号"] == style_no].copy()
        if style_df.empty:
            return pd.DataFrame()
        style_df[c["date"]] = pd.to_datetime(style_df[c["date"]], errors="coerce")
        style_df = style_df[(style_df[c["date"]] >= start) & (style_df[c["date"]] <= end)]
        if style_df.empty:
            return pd.DataFrame()
        if "未收货标记" in style_df.columns:
            style_df["未收货标记"] = style_df["未收货标记"].fillna(False)
            unreceived_flag = style_df.groupby(
                [c["product"], c["color"], c["order"]], dropna=False
            )["未收货标记"].max().reset_index(name="未收货标记")
        else:
            unreceived_flag = None
        grouped = style_df.groupby(
            [c["product"], c["color"], c["order"]], dropna=False, as_index=False
        ).agg({
            c["rolls"]: "sum",
            c["amount"]: "sum",
            cfg["order"]["qty"]: "first",
            cfg["order"]["received"]: "first",
            "_mix_parsed": "first",          # v4: 保留混搭解析标记（单组内一致）
            "_mix_sub_colors": "first",      # v4: 保留子色明细
        })
        if unreceived_flag is not None:
            grouped = grouped.merge(unreceived_flag, on=[c["product"], c["color"], c["order"]], how="left")
            grouped["未收货标记"] = grouped["未收货标记"].fillna(False)
        return grouped

    def calculate_cost(self, df2: pd.DataFrame, q_override: float = None) -> Dict:
        """计算单件布料成本（双版本输出）。

        业务口径（用户定义）：
          - 实际收货/实际匹数：仅收货数≠0 的记录；收货=0 的布料视为备布。
          - Q1（不计备布）= 实际收货 / 实际匹数
          - 备布折算 M = 未收货匹数 × Q1
          - Q2（计入备布）= (实际收货 + M) / (实际匹数 + 未收货匹数)
          - K1（不计备布）= 实际布料金额 / 实际匹数 / Q1
          - K2（计入备布）= 全部布料金额(含未收货) / 全部匹数(含未收货) / Q2

        返回：{Q1, K1, K1-s, K1-p, Q2, K2, K2-s, K2-p}
        """
        if df2.empty:
            return {"Q1": 0.0, "K1": 0.0, "K1-s": 0.0, "K1-p": 0.0,
                    "Q2": 0.0, "K2": 0.0, "K2-s": 0.0, "K2-p": 0.0}
        c = self.cfg["cols"]
        rolls, amount, received = c["fabric"]["rolls"], c["fabric"]["amount"], c["order"]["received"]

        if "未收货标记" in df2.columns:
            actual_mask = ~df2["未收货标记"].fillna(False)
        else:
            actual_mask = df2[received] > 0
        actual = df2[actual_mask]
        unreceived = df2[~actual_mask]

        if q_override is not None:
            Q1 = q_override
        else:
            M_a, N_a = actual[received].sum(), actual[rolls].sum()
            Q1 = M_a / N_a if N_a > 0 else 0.0

        N_unr = unreceived[rolls].sum() if not unreceived.empty else 0.0
        M_backup = N_unr * Q1 if Q1 > 0 else 0.0
        M_total = actual[received].sum() + M_backup
        N_total = actual[rolls].sum() + N_unr
        Q2 = M_total / N_total if N_total > 0 else Q1

        C_act, N_act = actual[amount].sum(), actual[rolls].sum()
        K1 = C_act / N_act / Q1 if N_act > 0 and Q1 > 0 else 0.0
        C_all, N_all = df2[amount].sum(), df2[rolls].sum()
        K2 = C_all / N_all / Q2 if N_all > 0 and Q2 > 0 else 0.0

        out = {"Q1": round(Q1, 0), "K1": round(K1, 2),
               "Q2": round(Q2, 0), "K2": round(K2, 2)}

        # v6: 混搭纳入印花开关 —— True 时所有混搭行强制归入印花桶（不管产品名称后缀）
        if self.treat_mixed_as_print and "_mix_parsed" in df2.columns:
            row_is_print = (df2["产品名称"].map(self._is_print)
                            | df2["_mix_parsed"].fillna(False))
        else:
            row_is_print = df2["产品名称"].map(self._is_print)

        for key_base, is_print in [("K1", False), ("K1", True), ("K2", False), ("K2", True)]:
            sub = df2[row_is_print == is_print]
            key = f"{key_base}-{'p' if is_print else 's'}"
            if sub.empty:
                out[key] = 0.0
                continue
            if "未收货标记" in sub.columns:
                sub_actual = sub[~sub["未收货标记"].fillna(False)]
            else:
                sub_actual = sub[sub[received] > 0]
            if key_base == "K1":
                Cs, Ns = sub_actual[amount].sum(), sub_actual[rolls].sum()
                Ms = sub_actual[received].sum()
                Qs = Ms / Ns if Ns > 0 else 0.0
                out[key] = round(Cs / Ns / Qs, 2) if Ns > 0 and Qs > 0 else 0.0
            else:
                sub_unr = sub[sub["未收货标记"].fillna(False)] if "未收货标记" in sub.columns else sub[sub[received] == 0]
                Cs, Ns = sub[amount].sum(), sub[rolls].sum()
                Ns_unr = sub_unr[rolls].sum() if not sub_unr.empty else 0.0
                Ms = sub_actual[received].sum() + Ns_unr * Q1
                Qs = Ms / Ns if Ns > 0 else 0.0
                out[key] = round(Cs / Ns / Qs, 2) if Ns > 0 and Qs > 0 else 0.0
        return out

    def allocate_mixed_received(self, df2: pd.DataFrame) -> pd.DataFrame:
        """对混搭行进行订单收货量的匹数比例分配。

        核心逻辑：
          1. 先统计每个订单号下所有布料行的子色匹数 → order_fabric_pools
          2. 对每行混搭色，按其在子色池中的匹数占比分配订单表的该颜色收货量
          3. 在混配色行上新增列 _allocated_recv（分配后的有效收货量）和 _effective_Q（用于 K 计算的 Q）

        参数：
          df2 - build_df2 输出的聚合数据（含 "_mix_parsed" 标记）

        返回：df2（已修改，新增 _allocated_recv / _effective_Q 列）
        """
        c = self.cfg["cols"]["fabric"]
        oc = self.cfg["cols"]["order"]

        # Step A: 预统计 — 构建 {订单号前缀: {子色名: 总匹数}} 的匹数池
        fabric_pool = {}
        for _, row in df2.iterrows():
            ord_num_clean = str(row[c["order"]]).replace(".", "").strip()
            if not row.get("_mix_parsed", False):
                continue
            sub_colors = row.get("_mix_sub_colors")
            if not sub_colors:
                continue
            total_rolls = row[c["rolls"]]
            mix_info = self._parse_mix_colors(row[c["color"]])
            if not mix_info:
                continue
            n_sub = mix_info["n_parts"]

            # 获取该行每个子色的权重（从「混搭比例」列或等比）
            ratio_col = "混搭比例"
            raw_ratio = row.get(ratio_col)
            weights = None
            if pd.notna(raw_ratio) and str(raw_ratio).strip():
                try:
                    vals = [float(x) for x in str(raw_ratio).split(":")]
                    weights = [v / sum(vals) for v in vals]
                except ValueError:
                    pass
            if weights is None:
                weights = [1.0 / n_sub] * n_sub

            for i, (sub_name, _) in enumerate(sub_colors):
                share_rolls = total_rolls * weights[i]
                if ord_num_clean not in fabric_pool:
                    fabric_pool[ord_num_clean] = {}
                fabric_pool[ord_num_clean][sub_name] = fabric_pool[ord_num_clean].get(sub_name, 0.0) + share_rolls

        # Step B: 对每行混配色进行收货量分配
        mixed_mask = df2["_mix_parsed"].fillna(False)
        if not mixed_mask.any():
            return df2

        for idx in df2[mixed_mask].index:
            row = df2.loc[idx]
            ord_num_clean = str(row[c["order"]]).replace(".", "").strip()
            sub_colors = row["_mix_sub_colors"]
            total_rolls = row[c["rolls"]]

            mix_info = self._parse_mix_colors(row[c["color"]])
            if not mix_info:
                continue
            n_sub = mix_info["n_parts"]

            # 再次获取权重
            ratio_col = "混搭比例"
            raw_ratio = row.get(ratio_col)
            weights = None
            if pd.notna(raw_ratio) and str(raw_ratio).strip():
                try:
                    vals = [float(x) for x in str(raw_ratio).split(":")]
                    weights = [v / sum(vals) for v in vals]
                except ValueError:
                    pass
            if weights is None:
                weights = [1.0 / n_sub] * n_sub

            allocated_recv_total = 0.0
            for i, (sub_name, _) in enumerate(sub_colors):
                weight_i = weights[i]
                share_rolls = total_rolls * weight_i

                # 在该订单号的该子色池中查找总匹数
                pool_sub = fabric_pool.get(ord_num_clean, {})
                total_pool_rolls = pool_sub.get(sub_name, 0.0)

                if total_pool_rolls <= 0:
                    continue

                ratio_in_pool = share_rolls / total_pool_rolls

                # 从订单表取该子色的总收货量（按订单编号精确匹配）
                order_rows = self.data["订单表"][
                    (self.data["订单表"][oc["order"]].astype(str).str.replace(".", "", regex=False).str.strip() == ord_num_clean) &
                    (self.data["订单表"][oc["color"]].astype(str).str.strip().str.contains(re.escape(sub_name), na=False))
                ]

                if not order_rows.empty:
                    total_order_recv = order_rows[oc["received"]].sum()
                    allocated_recv = total_order_recv * ratio_in_pool
                    allocated_recv_total += allocated_recv

            df2.at[idx, "_allocated_recv"] = allocated_recv_total
            if total_rolls > 0:
                df2.at[idx, "_effective_Q"] = allocated_recv_total / total_rolls
            else:
                df2.at[idx, "_effective_Q"] = 0.0

        return df2

    def calculate_mixed_costs(self, df2: pd.DataFrame, style_no: str, start, end) -> List[dict]:
        """对混合（混搭）行计算加权布料费用 K1/K2。

        算法：
          直接从 _mix_sub_colors 列获取订单级混配色列表（v5 修复）。
          不再用面料单色值去调用 _parse_mix_colors（那会失败，因为单色不含 "+"）。
          按匹数分组汇总后计算：同订单号的多条面料行属于同一混搭订单，
          先合并匹数/金额/收货量，再用子色池做匹数比例分配。

        返回：[{款号, 起始日期, 结束日期, 混合颜色, 匹数, 金额, 成衣数Q1, 布料费用K1-m, 混搭配比, 子色明细}, ...]
        """
        c = self.cfg["cols"]["fabric"]
        oc = self.cfg["cols"]["order"]
        cfg_all = self.cfg["cols"]       # v5: full cols dict for order key access
        result_rows = []

        mixed_df = df2[df2.get("_mix_parsed", False) == True].copy()
        if mixed_df.empty:
            return result_rows

        # Step 1: 按订单号分组汇总 — 同一订单号的多个面料行合并为一个混搭订单记录
        # v5 fix: df2 已经经过 build_df2 聚合，不再有 c["date"] 列，不需要转日期
        grouped_orders = mixed_df.groupby(c["order"]).agg({
            c["rolls"]: "sum",
            c["amount"]: "sum",
            cfg_all["order"]["received"]: "first",
            "_mix_sub_colors": "first",      # 子色列表在同一订单内一致
            c["color"]: lambda x: f"[{x.iloc[0]}+...] ({len(x)} 条面料行)",  # 展示首条 + 行数
        }).reset_index()
        grouped_orders.columns = [c["order"], c["rolls"], c["amount"],
                                   cfg_all["order"]["received"], "_mix_sub_colors",
                                   "_display_color"]

        # Step 2: 构建匹数池 — {订单号前缀: {子色名: 总匹数}}
        fabric_pool = {}
        for _, row in grouped_orders.iterrows():
            ord_num_clean = str(row[c["order"]]).replace(".", "").strip()
            sub_colors = row.get("_mix_sub_colors")
            if not sub_colors:
                continue
            total_rolls = row[c["rolls"]]
            mix_info = self._get_mix_info_from_sub_colors(sub_colors)
            if not mix_info:
                continue
            n_sub = mix_info["n_parts"]

            # 获取权重（从「混搭比例」列或等比）
            raw_ratio = row.get("混搭比例")
            weights = None
            if pd.notna(raw_ratio) and str(raw_ratio).strip():
                try:
                    vals = [float(x) for x in str(raw_ratio).split(":")]
                    weights = [v / sum(vals) for v in vals]
                except ValueError:
                    pass
            if weights is None:
                weights = [1.0 / n_sub] * n_sub

            for i, (sub_name, _) in enumerate(sub_colors):
                share_rolls = total_rolls * weights[i]
                if ord_num_clean not in fabric_pool:
                    fabric_pool[ord_num_clean] = {}
                fabric_pool[ord_num_clean][sub_name] = fabric_pool[ord_num_clean].get(sub_name, 0.0) + share_rolls

        # Step 3: 对每个混搭订单执行匹数比例分配 + 成本计算
        for _, row in grouped_orders.iterrows():
            ord_num_clean = str(row[c["order"]]).replace(".", "").strip()
            total_rolls = row[c["rolls"]]
            total_amt = row[c["amount"]]
            received_total = row[cfg_all["order"]["received"]]
            sub_colors = row.get("_mix_sub_colors")

            if not sub_colors:
                continue

            mix_info = self._get_mix_info_from_sub_colors(sub_colors)
            if not mix_info:
                continue

            n_sub = mix_info["n_parts"]
            weights = [1.0 / n_sub] * n_sub  # 兜底等比

            # 分配收货量
            allocated_recv_total = 0.0
            for i, (sub_name, _) in enumerate(sub_colors):
                weight_i = weights[i]
                share_rolls = total_rolls * weight_i

                pool_sub = fabric_pool.get(ord_num_clean, {})
                total_pool_rolls = pool_sub.get(sub_name, 0.0)

                if total_pool_rolls <= 0:
                    continue

                ratio_in_pool = share_rolls / total_pool_rolls
                order_rows = self.data["订单表"][
                    (self.data["订单表"][oc["order"]].astype(str).str.replace(".", "", regex=False).str.strip() == ord_num_clean) &
                    (self.data["订单表"][oc["color"]].astype(str).str.strip().str.contains(re.escape(sub_name), na=False))
                ]
                if not order_rows.empty:
                    total_order_recv = order_rows[oc["received"]].sum()
                    allocated_recv_total += total_order_recv * ratio_in_pool

            # 计算 Q1
            Q1 = allocated_recv_total / total_rolls if total_rolls > 0 else (received_total / total_rolls if total_rolls > 0 else 0.0)

            # 计算加权平均单价
            weighted_unit_sum = 0.0
            sub_details = []
            for i, (sub_name, is_print) in enumerate(zip([sc[0] for sc in sub_colors], [sc[1] for sc in sub_colors])):
                unit_price = total_amt / total_rolls if total_rolls > 0 else 0.0
                weighted_unit_sum += weights[i] * unit_price
                cat = "印花" if is_print else "纯色素"
                sub_details.append({"子色名称": sub_name, "类别": cat,
                                    "权重": round(weights[i], 4), "子色单价": round(unit_price, 2)})

            # K1 = 加权平均单价 ÷ Q1
            K1_m = weighted_unit_sum / Q1 if Q1 > 0 else 0.0
            K2_m = K1_m  # 简化：暂不设备布折算差异

            # 混合颜色字符串 — 简化为 "N子色" 格式（v6 改进）
            n_sub = len(sub_colors)
            n_print = sum(1 for _, ip in sub_colors if ip)
            n_solid = n_sub - n_print
            mix_display = f"{n_sub}子色"

            result_rows.append({
                "款号": style_no,
                "起始日期": start.strftime("%Y-%m-%d"),
                "结束日期": end.strftime("%Y-%m-%d"),
                "混合颜色": mix_display,
                "匹数": total_rolls,
                "金额": total_amt,
                "成衣数Q1": round(Q1, 0),
                "布料费用K1-m": round(K1_m, 2),
                "混搭配比": ": ".join(f"{round(w*100,1)}%" for w in weights),
                "子色明细": "; ".join(f"{sd['子色名称']}({sd['类别']}: {sd['权重']})" for sd in sub_details),
                "_n_mix_orders": 1,     # 该款式涉及的混搭订单数（后续汇总用）
            })

        return result_rows

    @staticmethod
    def _get_mix_info_from_sub_colors(sub_colors) -> Optional[dict]:
        """从已存在的 _mix_sub_colors 列创建混合信息（v5 新增辅助方法）。

        替代原来的 _parse_mix_colors，直接接受 [(name, is_print), ...] 格式。
        """
        if not sub_colors or len(sub_colors) < 2:
            return None
        return {"sub_colors": sub_colors, "n_parts": len(sub_colors)}

    def process_backup_fabric(self, df0: pd.DataFrame, q_values: Dict[str, float]) -> pd.DataFrame:
        if df0.empty:
            return df0
        c = self.cfg["cols"]
        df0 = df0.copy()
        df0["成衣数Q"] = df0[c["fabric"]["product"]].map(
            lambda p: q_values.get(self._extract_style_from_product_name(p), 0.0))
        rolls = pd.to_numeric(df0[c["fabric"]["rolls"]], errors="coerce").fillna(0)
        df0["已收货数量"] = df0["成衣数Q"] * rolls
        return df0

    def get_valid_style_numbers(self) -> List[str]:
        style_info = self.data["款号信息表"]
        pos_col = self.cfg["cols"]["style_info"]["position"]
        return style_info[style_info[pos_col] != self.cfg["eliminated_position"]][
            self.cfg["cols"]["style_info"]["style"]].tolist()

    # ---------- 模式 A：成衣数手填 ----------
    def _load_profile_q(self) -> Dict[str, float]:
        """从「款号档案」读取成衣数：{款号: 出货数（件/匹）}

        业务口径：出货数为空（或 0）= 该款已淘汰，不参与计算，不视为数据缺失。
        """
        pc = self.cfg["profile_q_cols"]
        profile = self.data["款号档案"]
        if pc["style"] not in profile.columns:
            raise ValueError(f"「款号档案」缺少列「{pc['style']}」")
        if pc["q"] not in profile.columns:
            raise ValueError(
                f"「款号档案」缺少成衣数列「{pc['q']}」。"
                f"请在款号档案中手工填写每匹布出多少件成衣（留空 = 该款已淘汰，不参与计算）。")

        q_map = {}
        skipped = []
        for style, q in zip(profile[pc["style"]], profile[pc["q"]]):
            if pd.isna(style):
                continue
            key = str(style).strip().upper()
            qq = pd.to_numeric(q, errors="coerce")
            if pd.notna(qq) and qq > 0:
                q_map[key] = float(qq)
            else:
                skipped.append(key)
        self.diag["款号档案总行数"] = len(profile)
        self.diag["有出货数(成衣数)的款"] = len(q_map)
        self.diag["款号档案中出货数为空(视为已淘汰)的款"] = skipped
        return q_map

    def _calculate_profile_q(self):
        """模式 A 主流程：K1 = Σ金额 ÷ Σ匹数 ÷ Q，纯色/印花分桶"""
        c = self.cfg["cols"]["fabric"]
        fabric = self.data["布料跟踪表"].copy()

        q_map = self._load_profile_q()
        if self.q_overrides:
            q_map.update({str(k).strip().upper(): float(v)
                          for k, v in self.q_overrides.items()})

        # --- 清洗 ---
        fabric["_date"] = pd.to_datetime(fabric[c["date"]], errors="coerce")
        n_raw = len(fabric)
        n_bad_date = int(fabric["_date"].isna().sum())
        # 基准日必须在「剔除未送货行」之前取：
        # 最新下的单往往匹数/金额仍为空，若用清洗后的最大日期，窗口终点会整体前移。
        raw_data_end = fabric["_date"].max()
        fabric = fabric.dropna(subset=["_date"]).copy()

        fabric["款号"] = fabric[c["product"]].map(self._extract_style_from_product_name)
        fabric["_rolls"] = pd.to_numeric(fabric[c["rolls"]], errors="coerce")
        fabric["_amount"] = pd.to_numeric(fabric[c["amount"]], errors="coerce")
        fabric["_is_print"] = fabric[c["product"]].map(self._is_print)

        n_no_style = int(fabric["款号"].isna().sum())
        # 送货匹数与金额必须有效（剔除「取消」等文本、0 匹 0 元行）
        mask_ok = (fabric["_rolls"].fillna(0) > 0) & (fabric["_amount"].fillna(0) > 0)
        n_dropped = int((~mask_ok).sum())
        fabric = fabric[mask_ok].copy()

        fabric_styles = set(fabric["款号"].dropna().unique())
        self.diag.update({
            "布料跟踪表总行数": n_raw,
            "下单日期无效行": n_bad_date,
            "款号无法识别行": n_no_style,
            "匹数或金额无效被剔除行": n_dropped,
            "有效行": len(fabric),
            "有布料数据的款": len(fabric_styles),
        })

        # --- 目标款：有布料数据 且 有 Q 值 ---
        targets = sorted(fabric_styles & set(q_map.keys()))
        if self.styles:
            targets = [s for s in targets if s in self.styles]
        no_q = sorted(fabric_styles - set(q_map.keys()))
        self.diag["有待算款(有布料且有Q)"] = len(targets)
        self.diag["有布料但出货数为空(已淘汰,不计算)"] = no_q

        # --- 统一时间窗口：基准日 = 布料跟踪表最新下单日（含尚未送货的行） ---
        data_end = raw_data_end
        default_days = self.cfg.get("default_lookback_days", 365)
        self.diag["数据基准日"] = data_end.strftime("%Y-%m-%d")

        def period_for(style):
            """默认统一近一年；periods 显式指定时按款覆盖。"""
            p = self.periods.get(style)
            if not p:
                return data_end - timedelta(days=default_days), data_end
            s = _parse_date(p.get("start"))
            e = _parse_date(p.get("end"))
            days = p.get("days")
            if s is None and e is not None and days:
                s = e - timedelta(days=days)
            if e is None and s is not None and days:
                e = s + timedelta(days=days)
            if e is None:
                e = data_end
            if s is None:
                s = e - timedelta(days=days or default_days)
            return s, e

        results = []
        for style in targets:
            start, end = period_for(style)
            sub = fabric[(fabric["款号"] == style) &
                         (fabric["_date"] >= start) &
                         (fabric["_date"] <= end)]
            if sub.empty:
                continue
            q = q_map[style]
            amt = float(sub["_amount"].sum())
            rolls = float(sub["_rolls"].sum())
            if rolls <= 0 or q <= 0:
                continue

            def bucket(mask):
                s2 = sub[mask]
                r2 = float(s2["_rolls"].sum())
                a2 = float(s2["_amount"].sum())
                return (round(a2 / r2 / q, 2) if r2 > 0 else 0.0), a2, r2

            k1 = amt / rolls / q
            k1s, amt_s, rolls_s = bucket(~sub["_is_print"])
            k1p, amt_p, rolls_p = bucket(sub["_is_print"])

            results.append({
                "款号": style,
                "起始日期": start.strftime("%Y-%m-%d"),
                "结束日期": end.strftime("%Y-%m-%d"),
                "成衣数Q1": round(q, 0),
                "布料费用K1": round(k1, 2),
                "纯色布料费用K1-s": k1s,
                "印花布料费用K1-p": k1p,
                "成衣数Q2": round(q, 0),
                "布料费用K2": round(k1, 2),
                "纯色布料费用K2-s": k1s,
                "印花布料费用K2-p": k1p,
                # 便于核对的分项（不影响网页主表）
                "布料金额": round(amt, 2),
                "布料匹数": round(rolls, 0),
                "纯色金额": round(amt_s, 2),
                "纯色匹数": round(rolls_s, 0),
                "印花金额": round(amt_p, 2),
                "印花匹数": round(rolls_p, 0),
                "记录数": len(sub),
            })
            self.period_summary.append({"款号": style,
                                        "起始日期": start.strftime("%Y-%m-%d"),
                                        "结束日期": end.strftime("%Y-%m-%d")})

        self.diag["实际出结果款数"] = len(results)
        # 有出货数、但时间窗口内没有布料下单记录的款（通常是已停产、只清库存）
        got = {r["款号"] for r in results}
        self.diag["有出货数但窗口内无布料数据"] = sorted(
            s for s in targets if s not in got)
        return results, pd.DataFrame()

    # ---------- 主流程 ----------
    def calculate(self):
        self.load_data()
        if self.use_profile_q:
            return self._calculate_profile_q()

        for sheet, cols in zip(self.cfg["sheets"],
                               [self.cfg["cols"]["fabric"], self.cfg["cols"]["order"],
                                self.cfg["cols"]["mapping"], self.cfg["cols"]["style_info"]]):
            self._require_cols(self.data[sheet], cols, sheet)

        df1 = self.build_df1()
        df1_1 = self.build_df1_1(df1)
        df1_2 = self.build_df1_2(df1_1)

        target_styles = self.get_valid_style_numbers()
        if self.styles:
            target_styles = [s for s in target_styles if s in self.styles]

        q_values = {}
        periods = {}
        for style_no in target_styles:          # 第一轮：解析时间段 + 算 Q1
            start, end = self.resolve_period(style_no)
            periods[style_no] = (start, end)
            if style_no in self.q_overrides:
                q_values[style_no] = float(self.q_overrides[style_no])
                continue
            df2 = self.build_df2(df1_2, style_no, start, end)
            if not df2.empty:
                q_values[style_no] = self.calculate_cost(df2)["Q1"]

        results = []                            # 第二轮：并入备布算最终 K1 + 混搭
        backup_all = []
        cfg = self.cfg["cols"]
        c = cfg["fabric"]
        for style_no in target_styles:
            start, end = periods[style_no]
            df2 = self.build_df2(df1_2, style_no, start, end)
            df0 = self.build_df0_prepare(style_no, start, end)
            df0 = self.process_backup_fabric(df0, q_values)
            if not df0.empty:
                backup_all.append(df0)

            # 未收货行按备布折算
            if not df2.empty and self.cfg.get("treat_unreceived_as_backup", False):
                q = q_values.get(style_no, 0)
                if q > 0:
                    unreceived = df2.get("未收货标记", pd.Series(False, index=df2.index)).fillna(False)
                    rolls = pd.to_numeric(df2[c["rolls"]], errors="coerce").fillna(0)
                    df2.loc[unreceived, cfg["order"]["received"]] = q * rolls[unreceived]

            if df2.empty:
                if not df0.empty and q_values.get(style_no, 0) > 0:
                    q = q_values[style_no]
                    amt, rolls = df0[c["amount"]].sum(), pd.to_numeric(
                        df0[c["rolls"]], errors="coerce").sum()
                    if rolls > 0:
                        results.append({"款号": style_no,
                                        "起始日期": start.strftime("%Y-%m-%d"),
                                        "结束日期": end.strftime("%Y-%m-%d"),
                                        "成衣数Q1": q, "布料费用K1": round(amt / rolls / q, 2),
                                        "纯色布料费用K1-s": 0.0, "印花布料费用K1-p": 0.0,
                                        "成衣数Q2": q, "布料费用K2": round(amt / rolls / q, 2),
                                        "纯色布料费用K2-s": 0.0, "印花布料费用K2-p": 0.0})
                continue

            # 先做混搭收货量分配（在所有处理之前，针对原始 df2）
            df2_with_allocation = self.allocate_mixed_received(df2)

            # v4 fix: 如果有混搭行，从常规计算的 df2 中剔除它们，避免重复计入
            # v6: treat_mixed_as_print=True 时不剔除 —— 混搭行留在常规管线，
            #     其子色（多为印花）自然计入 K1/K1-p，即"混搭纳入印花体系"
            has_mixed = (df2_with_allocation["_mix_parsed"].fillna(False)).any()
            if has_mixed and not self.treat_mixed_as_print:
                df2_regular = df2_with_allocation[~df2_with_allocation["_mix_parsed"].fillna(False)]
            else:
                df2_regular = df2_with_allocation

            # v6 fix: df2_all 必须以剔除混搭后的 df2_regular 为基础
            # （此前 df2_all = df2.copy()，在无备布数据时会把混搭行漏回常规计算）
            df2_all = df2_regular.copy()

            if not df0.empty:
                recs = pd.DataFrame({
                    c["product"]: df0[c["product"]],
                    c["color"]: df0[c["color"]],
                    c["order"]: df0[c["order"]].astype(str) + "_备布",
                    c["rolls"]: df0[c["rolls"]],
                    c["amount"]: df0[c["amount"]],
                    cfg["order"]["qty"]: 0,
                    cfg["order"]["received"]: df0["已收货数量"],
                    "未收货标记": True,
                })
                df2_all = pd.concat([df2_all, recs], ignore_index=True)

            # 计算混搭成本（如果该行是混搭色）
            # v6: 无论开关状态都计算并记录到 mixed_results_all（供"混搭明细"表展示
            #     和"布料费用结果"的混搭标识列使用）；但只有 treat_mixed_as_print=False
            #     时才把 K1-m 独立行并入最终结果
            mixed_rows = self.calculate_mixed_costs(df2_with_allocation, style_no, start, end)
            if mixed_rows:
                self.mixed_results_all.extend(mixed_rows)
                if not self.treat_mixed_as_print:
                    for mr in mixed_rows:
                        results.append(mr)

            cost = self.calculate_cost(df2_all, q_override=q_values.get(style_no, 0))
            results.append({"款号": style_no,
                            "起始日期": start.strftime("%Y-%m-%d"),
                            "结束日期": end.strftime("%Y-%m-%d"),
                            "成衣数Q1": cost["Q1"], "布料费用K1": cost["K1"],
                            "纯色布料费用K1-s": cost["K1-s"], "印花布料费用K1-p": cost["K1-p"],
                            "成衣数Q2": cost["Q2"], "布料费用K2": cost["K2"],
                            "纯色布料费用K2-s": cost["K2-s"], "印花布料费用K2-p": cost["K2-p"]})
            self.period_summary.append({"款号": style_no,
                                        "起始日期": start.strftime("%Y-%m-%d"),
                                        "结束日期": end.strftime("%Y-%m-%d")})

        backup_df = pd.concat(backup_all, ignore_index=True) if backup_all else pd.DataFrame()
        return results, backup_df


def run_fabric_calculation(uploaded_file, config: dict = None,
                           styles: Optional[List[str]] = None,
                           periods: Optional[Dict[str, dict]] = None,
                           q_overrides: Optional[Dict[str, float]] = None,
                           treat_mixed_as_print: bool = False,
                           use_profile_q: bool = False):
    """统一入口。
    参数：
      uploaded_file - Excel 路径或文件对象
      config        - 覆盖默认 CONFIG（如 default_lookback_days）
      styles        - 只计算这些款号；None=全部非淘汰款
      periods       - {款号: {"start":..,"end":..}} 或 {款号: {"days": N}} 自定义时间段
      q_overrides   - {款号: 成衣数} 人工核定成衣数（优先级高于款号档案）
      treat_mixed_as_print - 混搭纳入印花开关（仅订单表模式有效）
      use_profile_q - True=手填成衣数模式：只读「布料跟踪表」+「款号档案」，
                      Q 取款号档案「出货数（件/匹）」；出货数为空 = 已淘汰，直接跳过
    返回：
      {"布料费用结果", "备布数据", "时间段汇总", "数据质量报告", "字段说明", "混搭明细", "数据诊断"}
    """
    calculator = FabricCostCalculator(uploaded_file, config, styles=styles,
                                      periods=periods, q_overrides=q_overrides,
                                      treat_mixed_as_print=treat_mixed_as_print,
                                      use_profile_q=use_profile_q)
    results, df0 = calculator.calculate()

    # Separate regular and mixed results
    regular_results = [r for r in results if "混合颜色" not in r]
    mixed_results = calculator.mixed_results_all   # v6: 始终计算，供标识列和明细表使用

    df_results = pd.DataFrame(regular_results)
    if not df_results.empty:
        base_cols = ["款号", "起始日期", "结束日期", "成衣数Q1",
                     "布料费用K1", "纯色布料费用K1-s", "印花布料费用K1-p",
                     "成衣数Q2", "布料费用K2", "纯色布料费用K2-s", "印花布料费用K2-p"]
        if use_profile_q:
            # 手填模式保留核对用的分项金额/匹数
            extra = ["布料金额", "布料匹数", "纯色金额", "纯色匹数",
                     "印花金额", "印花匹数", "记录数"]
            base_cols += [cc for cc in extra if cc in df_results.columns]
        df_results = df_results[base_cols]

    if not df_results.empty and not use_profile_q:
        # ---------- v6: 追加混搭专用列（有混搭的款填值，无混搭的留空/否） ----------
        mixed_summary = {}
        for mr in mixed_results:
            sn = mr["款号"]
            if sn not in mixed_summary:
                mixed_summary[sn] = {"orders": 0, "rolls": 0.0, "amt": 0.0,
                                     "Q1": None, "K1m": None, "detail": ""}
            ms = mixed_summary[sn]
            ms["orders"] += 1
            ms["rolls"] += float(mr.get("匹数", 0) or 0)
            ms["amt"] += float(mr.get("金额", 0) or 0)
            if ms["Q1"] is None:
                ms["Q1"] = mr.get("成衣数Q1")
                ms["K1m"] = mr.get("布料费用K1-m")
                ms["detail"] = mr.get("子色明细", "")

        flag = "是（已并入印花）" if treat_mixed_as_print else "是"
        df_results["是否混搭"] = df_results["款号"].map(
            lambda s: flag if s in mixed_summary else "否")
        df_results["混搭订单数"] = df_results["款号"].map(
            lambda s: mixed_summary[s]["orders"] if s in mixed_summary else None)
        df_results["混搭匹数"] = df_results["款号"].map(
            lambda s: round(mixed_summary[s]["rolls"], 1) if s in mixed_summary else None)
        df_results["混搭金额"] = df_results["款号"].map(
            lambda s: round(mixed_summary[s]["amt"], 2) if s in mixed_summary else None)
        df_results["混搭成衣数Q"] = df_results["款号"].map(
            lambda s: mixed_summary[s]["Q1"] if s in mixed_summary else None)
        # 开关开启时混搭已并入印花（K1-p 已包含），不再展示独立 K1-m 数值
        df_results["布料费用K1-m"] = df_results["款号"].map(
            lambda s: (mixed_summary[s]["K1m"] if s in mixed_summary else None)
            if not treat_mixed_as_print else None)
        df_results["混搭子色明细"] = df_results["款号"].map(
            lambda s: mixed_summary[s]["detail"] if s in mixed_summary else None)
        df_results["混搭订单数"] = df_results["混搭订单数"].astype("Int64")  # 整数显示（无混搭为 <NA>）

    out_cols = ["订单编号", "产品名称", "颜色&色号", "布料送货数量（匹）", "金额（元）",
                "成衣数Q", "已收货数量", "备布标记"]
    df0_out = df0[[col for col in out_cols if col in df0.columns]] if not df0.empty else pd.DataFrame(columns=out_cols)

    mixed_cols = ["款号", "起始日期", "结束日期", "混合颜色", "匹数", "金额",
                  "成衣数Q1", "布料费用K1-m", "混搭配比", "子色明细"]
    df_mixed = pd.DataFrame(mixed_results)
    if not df_mixed.empty:
        df_mixed = df_mixed[[c for c in mixed_cols if c in df_mixed.columns]]

    return {
        "布料费用结果": df_results,
        "备布数据": df0_out,
        "时间段汇总": pd.DataFrame(calculator.period_summary),
        "数据质量报告": pd.DataFrame(calculator.quality_report),
        "字段说明": pd.DataFrame(FIELD_GLOSSARY),
        "混搭明细": df_mixed,
        "数据诊断": calculator.diag,
    }


# ============ 输出字段说明（分享给不熟悉输出格式的人查看） ============
FIELD_GLOSSARY = [
    {"字段": "（手填模式）成衣数来源", "类型": "口径",
     "含义": "v7 起成衣数由业务在「款号档案 - 出货数（件/匹）」手工填写，不再由订单表推算",
     "计算口径": "有填写的款才计算；出货数为空的款视为已淘汰，直接跳过，不参与结果"},
    {"字段": "款号", "类型": "标识", "含义": "服装款号，如 PPY0026（每个款独立计算）",
     "计算口径": "-"},
    {"字段": "起始日期 / 结束日期", "类型": "时间段", "含义": "本次计算覆盖的时间范围",
     "计算口径": "默认近一年（365 天）；可用 periods 参数按款自定义"},
    {"字段": "成衣数Q1（不计备布）", "类型": "成衣数", "含义": "每匹布实际能做出的成衣件数",
     "计算口径": "Q1 = 实际收货件数 ÷ 实际布料匹数（只统计已收货>0 的记录；未收货/备布的匹数不计入分母）；可被人工核定值覆盖"},
    {"字段": "布料费用K1（不计备布）", "类型": "成本", "含义": "每件成衣分摊的布料成本（不含未收货的布）",
     "计算口径": "K1 = 实际布料金额 ÷ 实际匹数 ÷ Q1（等价于：实际布料金额 ÷ 实际成衣数）"},
    {"字段": "纯色布料费用K1-s", "类型": "成本", "含义": "纯色布部分每件成衣的布料成本",
     "计算口径": "用该款纯色布自己的 金额/匹数/实际收货 单独计算"},
    {"字段": "印花布料费用K1-p", "类型": "成本", "含义": "印花布部分每件成衣的布料成本",
     "计算口径": "用该款印花布自己的 金额/匹数/实际收货 单独计算"},
    {"字段": "成衣数Q2（计入备布）", "类型": "成衣数", "含义": "把未收货（假定备布）的布也按 Q1 折算成衣数后的每匹产出",
     "计算口径": "Q2 = (实际收货 + 未收货匹数×Q1) ÷ (实际匹数 + 未收货匹数)"},
    {"字段": "布料费用K2（计入备布）", "类型": "成本", "含义": "全部布料（含未收货/备布）都计入后每件成衣的布料成本",
     "计算口径": "K2 = 全部布料金额 ÷ 全部匹数 ÷ Q2（= 总布料金额 ÷ 折算后总成衣数）"},
    {"字段": "纯色布料费用K2-s / 印花布料费用K2-p", "类型": "成本", "含义": "K2 的纯色 / 印花拆分",
     "计算口径": "各用自己子集的全部金额/匹数，未收货部分按 Q1 折算"},
    {"字段": "（附）未收货 / 备布", "类型": "概念", "含义": "已经买了布料、但对应成衣订单还未收货（或订单表查不到）的布料",
     "计算口径": "这类布按'假定能做出与已收货布相同件数'折算，不直接丢弃"},
    {"字段": "（附）数据质量报告", "类型": "概念", "含义": "源数据重复行检查结果（订单表按 订单+SKU、布料表全列、映射表按SKU）",
     "计算口径": "删除重复行>0 时说明源数据有整单重复，需向上游反馈修正"},
    {"字段": "是否混搭", "类型": "标识", "含义": "该款在时间窗口内是否存在混搭订单（订单表颜色名含'+'）",
     "计算口径": "是=独立计算模式；是（已并入印花）=开关勾选后混搭已计入K1-p；否=无混搭"},
    {"字段": "混搭订单数 / 混搭匹数 / 混搭金额 / 混搭成衣数Q", "类型": "混搭汇总",
     "含义": "该款混搭订单的规模汇总：涉及几个订单、共多少匹布、总金额、按匹数比例分配后的成衣数",
     "计算口径": "混搭成衣数Q = 订单收货量按子色匹数占比分配后 ÷ 混搭匹数"},
    {"字段": "布料费用K1-m", "类型": "成本", "含义": "混搭部分每件成衣的布料成本（仅独立计算模式有值）",
     "计算口径": "加权平均单价（等比或按'混搭比例'列）÷ 混搭成衣数Q；开关勾选时已并入K1-p，此列为空"},
    {"字段": "混搭子色明细", "类型": "混搭明细", "含义": "该款混搭包含的子色清单：名称(类别: 权重)",
     "计算口径": "子色名取自订单表颜色名按'+'拆分；含YH/GG编码的判定为印花，其余为纯色素"},
]


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python fabric_cost.py <Excel路径> [款号1,款号2...]")
        sys.exit(1)
    styles = sys.argv[2].split(",") if len(sys.argv) > 2 else None
    out = run_fabric_calculation(sys.argv[1], styles=styles)
    print(out["布料费用结果"].to_string(index=False))
    print("\n[时间段汇总]")
    print(out["时间段汇总"].to_string(index=False))
