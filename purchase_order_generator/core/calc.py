# -*- coding: utf-8 -*-
"""calc.py — 计算模块：比例/条数/尺码/类型/混搭色工具函数"""
import re
from math import gcd
from functools import reduce
from typing import Dict, List, Tuple

# 尺码标准顺序（上限 S~6XL）
SIZE_ORDER = ["S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL"]

# 布行分隔符：逗号(中英文)/分号(中英文)/空格
SUPPLIER_SEP_RE = re.compile(r"[,;，；\s]+")


def split_suppliers(supplier: str) -> List[str]:
    """按 逗号/分号/空格 拆分布行，返回非空部分列表。
    例: '剑邑布行,宏裕布行' / '剑邑布行 宏裕布行' / '剑邑布行；宏裕布行' -> ['剑邑布行','宏裕布行']"""
    if not supplier:
        return []
    return [p.strip() for p in SUPPLIER_SEP_RE.split(str(supplier)) if p.strip()]


def sort_sizes(sizes: List[str]) -> List[str]:
    """按 S~6XL 标准顺序排序尺码列表"""
    rank = {s: i for i, s in enumerate(SIZE_ORDER)}
    return sorted(sizes, key=lambda s: rank.get(s, 99))


def gcd_ratio(qty_by_size: Dict[str, int]) -> str:
    """按各尺码数量算 GCD 比例，如 {L:120, XL:200} -> '3:5'"""
    qs = [q for q in qty_by_size.values() if q and q > 0]
    if not qs:
        return ""
    g = reduce(gcd, qs)
    ordered = sort_sizes([s for s in qty_by_size if qty_by_size.get(s)])
    return ":".join(str(qty_by_size[s] // g) for s in ordered)


def rolls(qty: int, pcs_per_roll: int) -> int:
    """布料条数 = 件数 ÷ 每条布出货数，向上取整"""
    if not qty or qty <= 0:
        return 0
    if not pcs_per_roll or pcs_per_roll <= 0:
        pcs_per_roll = 40
    return -(-qty // pcs_per_roll)  # ceil division


def is_mixed_color(color: str) -> bool:
    """混搭色判定：颜色名含 +（如 #179 海军蓝+YH5126 棕蓝格子）"""
    return bool(color) and "+" in str(color)


def split_mixed_color(color: str) -> Tuple[str, str]:
    """混搭色按第一个 + 拆成 (前段, 后段)"""
    parts = str(color).split("+", 1)
    return parts[0].strip(), parts[1].strip() if len(parts) > 1 else ""


def half_split(qty: int) -> Tuple[int, int]:
    """对半分，奇数时前段多 1（如 51 -> (26, 25)）"""
    half = qty // 2
    return qty - half, half


def classify_type(sku_names: List[str]) -> str:
    """类型拼接：按 SKU 名称关键词（纯色/印花/混搭），混搭算印花。
    例: 只有纯色 -> '纯色'; 纯色+印花 -> '纯色 印花'; 只有混搭 -> '印花'"""
    has_chunse, has_yinhua = False, False
    for name in sku_names:
        n = str(name)
        if "混搭" in n or "印花" in n:
            has_yinhua = True
        if "纯色" in n:
            has_chunse = True
    parts = []
    if has_chunse:
        parts.append("纯色")
    if has_yinhua:
        parts.append("印花")
    return " ".join(parts) if parts else ""


def color_kind(color: str) -> str:
    """颜色类型：混搭/纯色/印花（供分组、判断用）
    - 含 '+' -> 混搭
    - 以 YH/HX/... 印花编号开头或含 YH/HX 编号 -> 印花
    - 以 # 开头 -> 纯色
    - 兜底：含 'YH' 或 'HX' 视为印花，否则纯色"""
    c = str(color).strip()
    if "+" in c:
        return "混搭"
    first = c.split()[0] if c.split() else c
    if first.startswith("#"):
        return "纯色"
    return "印花"


if __name__ == "__main__":
    # 自测
    assert gcd_ratio({"L": 120, "XL": 200, "2XL": 880, "3XL": 1000, "4XL": 400, "5XL": 400}) == "3:5:22:25:10:10"
    assert rolls(50, 40) == 2 and rolls(40, 40) == 1 and rolls(39, 40) == 1
    assert split_mixed_color("#179 海军蓝+YH5126 棕蓝格子") == ("#179 海军蓝", "YH5126 棕蓝格子")
    assert half_split(51) == (26, 25)
    assert classify_type(["PPY4005-YH 亨利领睡衣套装 纯色", "PPY4005-YH 亨利领睡衣套装 印花"]) == "纯色 印花"
    assert classify_type(["PPY4005-YH 亨利领睡衣套装 混搭"]) == "印花"
    assert classify_type(["PPY4005-CS 亨利领睡衣套装 纯色"]) == "纯色"
    assert sort_sizes(["2XL", "S", "5XL", "M"]) == ["S", "M", "2XL", "5XL"]
    assert color_kind("#179 海军蓝") == "纯色"
    assert color_kind("YH5126 棕蓝格子") == "印花"
    assert color_kind("#179 海军蓝+YH5126 棕蓝格子") == "混搭"
    # 布行分隔符
    assert split_suppliers("剑邑布行,宏裕布行") == ["剑邑布行", "宏裕布行"]
    assert split_suppliers("剑邑布行 宏裕布行") == ["剑邑布行", "宏裕布行"]
    assert split_suppliers("剑邑布行；宏裕布行") == ["剑邑布行", "宏裕布行"]
    assert split_suppliers("剑邑布行，宏裕布行;云悦鸿") == ["剑邑布行", "宏裕布行", "云悦鸿"]
    assert split_suppliers("剑邑布行") == ["剑邑布行"]
    assert split_suppliers("") == []
    print("calc.py 自测全部通过 ✅")
