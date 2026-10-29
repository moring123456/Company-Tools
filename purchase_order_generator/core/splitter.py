# -*- coding: utf-8 -*-
"""splitter.py — 拆单逻辑

按 款号×工厂 组合拆成多份订单单元；每单元含:
  - 订单颜色行（颜色×尺码矩阵, 比例/总计/条数/备注）
  - 布料申购单分组（按布行; 混搭色按 SKU 行拆两半分别归两个布行）
订单号 = PPY + 输入文件名日期 + 序号(按组合首次出现顺序)
"""
import re
from dataclasses import dataclass, field
from typing import Dict, List, Tuple
from datetime import date

from .calc import (SIZE_ORDER, classify_type, gcd_ratio, half_split,
                   is_mixed_color, rolls, sort_sizes, split_mixed_color,
                   split_suppliers)
from .parser import InputData, SKURow, match_factory


# ---------- 数据结构 ----------

@dataclass
class OrderColorRow:
    color: str
    qty_by_size: Dict[str, int]      # 尺码→数量
    ratio: str = ""
    total: int = 0
    rolls: int = 0                   # 布料条数 = 各布行条数合计
    remark: str = ""
    bag_spec: str = ""
    is_mixed: bool = False


@dataclass
class SupplierSheetRow:
    color: str
    qty: int                         # 该布行该颜色件数
    rolls: int                       # 条数 = 件数÷pcs 向上取整
    supplier: str


@dataclass
class SupplierGroup:
    supplier: str
    fabric: str                      # 款级品名（主表 I 列，直接写入布料申购单「品名」）
    weight: str = ""                 # 重量（保留字段，暂不输出）
    rows: List[SupplierSheetRow] = field(default_factory=list)

    @property
    def total_rolls(self) -> int:
        return sum(r.rolls for r in self.rows)


@dataclass
class OrderUnit:
    order_no: str
    style: str
    product_name: str
    pattern: str                     # 纸样名称
    wash_label: str
    factory: str
    fabric: str                      # 款级品名
    need_bag: bool
    bag_spec: str
    type_str: str
    date_str: str = ""               # 文件名中的 8 位日期段 (从输入文件名提取, ui 可编辑)
    wash_components: List[str] = field(default_factory=list)   # 洗水唛成分列表（按行）
    sizes: List[str] = field(default_factory=list)
    color_rows: List[OrderColorRow] = field(default_factory=list)
    supplier_groups: List[SupplierGroup] = field(default_factory=list)
    total_qty: int = 0
    file_name: str = ""
    # v32: 工厂信息（配置表「工厂信息表」按简称匹配全称）
    factory_full: str = ""           # 工厂全称，如 嘉品服饰
    factory_contact: str = ""        # 联系人
    factory_phone: str = ""          # 联系电话
    factory_address: str = ""        # 工厂地址
    # v32: SKU 明细（送厂送货清单预填用）
    sku_items: List[Dict] = field(default_factory=list)   # [{sku,color,size,qty}]


@dataclass
class SplitResult:
    units: List[OrderUnit] = field(default_factory=list)
    mixed_splits: List[Dict] = field(default_factory=list)   # 混搭拆分明细(预览用)


# ---------- 工具 ----------

def _extract_date(filename: str) -> str:
    """从输入文件名提取 8 位日期 (PPY采购-20260827.xlsx -> 20260827)，失败用今天"""
    m = re.search(r"(\d{8})", filename)
    if m:
        return m.group(1)
    return date.today().strftime("%Y%m%d")


def _extract_style_prefix(style: str) -> str:
    """从款号提取字母前缀作为订单号前缀 (ADM1003 -> ADM, PPY4005 -> PPY, B0G51WSZ4Z -> B)"""
    m = re.match(r"^([A-Za-z]+)", style)
    if m:
        return m.group(1).upper()
    return "PPY"  # 兜底


def _supplier_parts(sr: SKURow) -> Tuple[str, str]:
    """混搭色布行拆两个（支持 逗号/空格/分号 分隔）；非混搭返回 (布行, '')"""
    if is_mixed_color(sr.color):
        parts = split_suppliers(sr.supplier)
        if len(parts) >= 2:
            return parts[0], parts[1]
    return sr.supplier, ""


# ---------- 拆单 ----------

def split(data: InputData, mixed_overrides: Dict[str, int] = None) -> SplitResult:
    """mixed_overrides: {sku: qty1} 混搭色前半段自定义件数（默认对半）"""
    mixed_overrides = mixed_overrides or {}
    result = SplitResult()
    date_str = _extract_date(data.input_filename)

    # 1. 按 (款号, 工厂) 分组，保持首次出现顺序
    order: List[Tuple[str, str]] = []
    groups: Dict[Tuple[str, str], List[SKURow]] = {}
    for sr in data.rows:
        key = (sr.style, sr.factory)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(sr)

    # 2. 订单号序号
    seq_by_key = {k: i + 1 for i, k in enumerate(order)}

    # 2. 订单号前缀：从款号提取字母部分（ADM1003 -> ADM, PPY4005 -> PPY）
    prefix_by_style: Dict[str, str] = {}

    for idx, key in enumerate(order):
        style, factory = key
        skus = groups[key]
        fi = data.fabrics.get(style)

        if style not in prefix_by_style:
            prefix_by_style[style] = _extract_style_prefix(style)
        prefix = prefix_by_style[style]

        unit = OrderUnit(
            order_no=f"{prefix}{date_str}{seq_by_key[key]:02d}",
            style=style,
            product_name=fi.product_name if fi else "",
            pattern=fi.pattern if fi else "",
            wash_label=fi.wash_label if fi else "",
            wash_components=fi.wash_components if fi else [],
            factory=factory,
            fabric=fi.fabric if fi else (skus[0].fabric if skus else ""),
            need_bag=fi.need_bag if fi else False,
            bag_spec=fi.bag_spec if fi else "",
            type_str=classify_type([s.sku_name for s in skus]),
            date_str=date_str,
        )
        color_qty: Dict[str, Dict[str, int]] = {}
        color_remark: Dict[str, set] = {}
        for sr in skus:
            color_qty.setdefault(sr.color, {})
            if sr.size:
                color_qty[sr.color][sr.size] = color_qty[sr.color].get(sr.size, 0) + max(sr.qty, 0)
            if sr.remark:
                color_remark.setdefault(sr.color, set()).add(sr.remark)

        # 4. 申购条目（按 SKU 行拆，混搭拆两半）
        #    entry = (supplier, color, qty, original_color)
        entries: List[Tuple[str, str, int, str]] = []
        # v32: 布行 → 重量（可选，保留不输出）
        supplier_meta: Dict[str, List[str]] = {}   # supplier -> [weight]

        def _collect_meta(sup: str, sr: SKURow):
            if not sup:
                return
            if sup not in supplier_meta:
                supplier_meta[sup] = [""]
            cur = supplier_meta[sup]
            if not cur[0] and sr.weight:
                cur[0] = sr.weight

        for sr in skus:
            q = max(sr.qty, 0)
            if is_mixed_color(sr.color):
                s1, s2 = _supplier_parts(sr)
                q1 = mixed_overrides.get(sr.sku)
                if q1 is None:
                    q1, q2 = half_split(q)
                else:
                    q1 = int(q1)
                    q2 = q - q1
                p1, p2 = split_mixed_color(sr.color)
                entries.append((s1, p1, q1, sr.color))
                _collect_meta(s1, sr)
                if s2 and q2:
                    entries.append((s2, p2, q2, sr.color))
                    _collect_meta(s2, sr)
                result.mixed_splits.append({
                    "order_no": unit.order_no, "style": style, "factory": factory,
                    "original_color": sr.color, "sku": sr.sku,
                    "part1": p1, "part2": p2,
                    "qty1": q1, "qty2": q2,
                    "supplier1": s1, "supplier2": s2,
                })
            else:
                entries.append((sr.supplier, sr.color, q, sr.color))
                _collect_meta(sr.supplier, sr)

        # 5. 布行聚合
        supplier_agg: Dict[str, Dict[str, int]] = {}
        for sup, color, q, _orig in entries:
            supplier_agg.setdefault(sup, {})
            supplier_agg[sup][color] = supplier_agg[sup].get(color, 0) + q

        pcs = fi.pcs_per_roll if fi else 40

        # 6. 生成申购单分组
        for sup in sorted(supplier_agg.keys(), key=lambda x: list(supplier_agg.keys()).index(x)):
            wt = supplier_meta.get(sup, [""])[0]
            group = SupplierGroup(supplier=sup, fabric=unit.fabric, weight=wt)
            for color, q in supplier_agg[sup].items():
                group.rows.append(SupplierSheetRow(color=color, qty=q, rolls=rolls(q, pcs), supplier=sup))
            unit.supplier_groups.append(group)

        # 7. 生成订单颜色行（条数 = 各布行该颜色聚合后条数合计，保证文件内自洽）
        color_orig: Dict[Tuple[str, str], str] = {}   # (supplier, color) -> 原色
        for sup, color, _q, orig in entries:
            color_orig[(sup, color)] = orig
        color_rolls: Dict[str, int] = {}
        for sup, agg in supplier_agg.items():
            for color, q in agg.items():
                orig = color_orig.get((sup, color), color)
                color_rolls[orig] = color_rolls.get(orig, 0) + rolls(q, pcs)

        # 记录输入顺序用于稳定排序（订单数相同时保持输入顺序）
        input_order = {color: i for i, color in enumerate(color_qty.keys())}
        for color in color_qty:
            qbs = color_qty[color]
            total = sum(qbs.values())
            c_rolls = color_rolls.get(color, 0)
            unit.color_rows.append(OrderColorRow(
                color=color,
                qty_by_size=qbs,
                ratio=gcd_ratio(qbs),
                total=total,
                rolls=c_rolls,
                remark=", ".join(sorted(color_remark.get(color, set()))),
                bag_spec=unit.bag_spec if unit.need_bag else "",
                is_mixed=is_mixed_color(color),
            ))
        # 按订单数倒序排（同件数按输入顺序稳定排序）
        unit.color_rows.sort(key=lambda cr: (-cr.total, input_order.get(cr.color, 0)))

        # 尺码固定为 S~6XL 全部 9 列（用户要求写死，不随数据变动）
        unit.sizes = list(SIZE_ORDER)

        # v32: 工厂信息（配置表「工厂信息表」按简称匹配全称，供申购单/送货清单使用）
        f_info = match_factory(data.config, factory)
        unit.factory_full = f_info.name or factory
        unit.factory_contact = f_info.contact
        unit.factory_phone = f_info.phone
        unit.factory_address = f_info.address

        # v32: SKU 明细（工厂送货清单预填：颜色按订单行顺序、尺码按 S~6XL 顺序）
        size_idx = {s: i for i, s in enumerate(SIZE_ORDER)}
        color_idx = {cr.color: i for i, cr in enumerate(unit.color_rows)}
        ordered = sorted(skus, key=lambda s: (color_idx.get(s.color, 999), size_idx.get(s.size, 99)))
        unit.sku_items = [
            {"sku": s.sku, "color": s.color, "size": s.size, "qty": max(s.qty, 0)}
            for s in ordered if max(s.qty, 0) > 0
        ]
        unit.total_qty = sum(cr.total for cr in unit.color_rows)
        unit.file_name = f"{unit.order_no}-{style} {unit.product_name} {unit.type_str} -{factory}.xlsx".replace("  ", " ")
        result.units.append(unit)

    return result
