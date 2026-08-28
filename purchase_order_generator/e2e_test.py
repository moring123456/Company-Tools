# -*- coding: utf-8 -*-
"""端到端测试: 解析 → 拆单 → 生成"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.parser import parse_input
from core.splitter import split
from core.generator import generate_all, make_zip

BASE = r"D:\WorkBuddy-存储\2026-06-24-11-13-53\purchase-order-generator"
input_path = os.path.join(BASE, "output", "PPY采购-20260827_测试输入.xlsx")
out_dir = os.path.join(BASE, "output", "generated")

data = parse_input(input_path)
print("=== 校验错误 ===")
if data.errors:
    for e in data.errors:
        print(" ", e)
else:
    print("  无错误 ✅")
print(f"SKU 行数: {len(data.rows)} | 面料表款数: {len(data.fabrics)} | 配置: {data.config}")

res = split(data)
print()
print("=== 拆单结果 ===")
for u in res.units:
    print(f"{u.order_no} | {u.style} | {u.factory} | 类型[{u.type_str}] | 尺码{u.sizes} | 件数{u.total_qty} | {u.file_name}")
    for cr in u.color_rows:
        print(f"    订单行: {cr.color} | {cr.qty_by_size} | 比例[{cr.ratio}] | 总计{cr.total} | 条数{cr.rolls} | 备注[{cr.remark}]")
    for g in u.supplier_groups:
        print(f"    申购单[{g.supplier}]: " + "; ".join(f"{r.color}={r.qty}件/{r.rolls}条" for r in g.rows))

print()
print("=== 混搭拆分明细 ===")
for m in res.mixed_splits:
    print(f"  {m['order_no']} {m['original_color']} -> [{m['part1']}@{m['supplier1']}:{m['qty1']}] + [{m['part2']}@{m['supplier2']}:{m['qty2']}]")

print()
print("=== 生成输出 ===")
paths = generate_all(data, res, out_dir)
for p in paths:
    print("  ", os.path.basename(p))
print("  汇总表:", os.path.basename(build_path := os.path.join(out_dir, f"{res.units[0].order_no[3:11]}_采购单汇总.xlsx")) if os.path.exists(os.path.join(out_dir, f"{res.units[0].order_no[3:11]}_采购单汇总.xlsx")) else "缺失")
zip_path = make_zip(out_dir, os.path.join(BASE, "output", f"PPY采购-20260827_采购单.zip"), res.units[0].order_no[3:11])
print("  zip:", os.path.basename(zip_path), os.path.getsize(zip_path), "bytes")
