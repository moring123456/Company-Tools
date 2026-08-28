# -*- coding: utf-8 -*-
"""generator.py — 输出文件生成

以模板为底版:
  - 订单 sheet:  复制 THD 模板「订单」sheet（保留样式/图片/底部固定内容），重写头部+主体
  - 布料申购单:  复制 PPY 旧模板「布料申购单」sheet，每个布行一张，删除右块，填数据
  - 汇总明细表:  独立生成 文件汇总 + 布行明细
"""
import os
import shutil
import zipfile
from copy import copy
from typing import Dict, List, Tuple

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.drawing.image import Image as XLImage

from .splitter import OrderUnit, SplitResult
from .parser import InputData

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
THD_TEMPLATE = os.path.join(BASE_DIR, "templates", "thd_template_clean.xlsx")  # 净化模板: 图片已内置
PPY_TEMPLATE = os.path.join(BASE_DIR, "templates", "ppy_old_template.xlsx")
IMAGES_DIR = os.path.join(BASE_DIR, "templates", "images")

THIN = Side(style="thin", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 模板内嵌图片（DISPIMG）→ 普通图片插回：位置名 / 默认图片ID / 目标紧凑尺寸(px)
# 比 THD 模板原 DISPIMG ext 更小（模板原图过大导致超出视觉区域）
DISPIMG_DEFAULTS = {
    "包装袋图片": ("ID_BE270A7E788F4E49844F73A78120EB23", 350, 230),
    "主唛": ("ID_6F3DBA6A97AB49F3AC43052A1ECB3BFC", 320, 240),
    "吊牌": ("ID_9C0A6E5FB0764A84BB59F70FBFA53CD6", 260, 110),
    "温馨提示卡片": ("ID_2A34054940634D1BAF0639E8B12A910E", 300, 250),
}


def _locate_image_anchors(ws) -> Dict[str, Tuple[int, int]]:
    """按标签行动态定位图片锚点（颜色行多时固定内容下移也能找到）
    规则: '包装袋图片'(A列标签) -> 图在其下一行 A 列; '主唛'(A列) -> 下一行 A 列;
          '吊牌'(D列标签) -> 下一行 D 列; '温馨提示'(B列) -> 本行 C 列"""
    anchors: Dict[str, Tuple[int, int]] = {}
    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        d = ws.cell(row=r, column=4).value
        b = ws.cell(row=r, column=2).value
        if isinstance(a, str) and "包装袋图片" in a:
            anchors["包装袋图片"] = (r + 1, 1)
        if isinstance(a, str) and "主唛" in a and "吊牌" not in a:
            anchors["主唛"] = (r + 1, 1)
        if isinstance(d, str) and "吊牌" in d:
            anchors["吊牌"] = (r + 1, 4)
        if b is not None and "温馨提示" in str(b):
            anchors["温馨提示卡片"] = (r, 3)
    return anchors


def _shift_images_down(ws, start_1based: int, extra: int):
    """插行后把锚点位于插行位置及其下方的图片整体下移 extra 行（0-based 处理）"""
    if extra <= 0:
        return
    for img in ws._images:
        try:
            r0 = img.anchor._from.row
        except (AttributeError, TypeError):
            continue
        if r0 >= start_1based - 1:
            img.anchor._from.row = r0 + extra


def _replace_cell_images(ws, unit_images: Dict[str, bytes] = None):
    """仅当输入文件/网页提供了替换图时，替换模板中对应位置的图片。
    模板里的默认图保留不动（方案2: 图已内置在净化模板中）"""
    import io
    unit_images = unit_images or {}
    if not unit_images:
        return
    anchors = _locate_image_anchors(ws)
    # 1. 删除模板中"有替换图"位置的旧图
    for img in list(ws._images):
        try:
            r0, c0 = img.anchor._from.row, img.anchor._from.col
        except (AttributeError, TypeError):
            continue
        for label, (row, col) in anchors.items():
            if r0 == row - 1 and c0 == col - 1 and label in unit_images:
                ws._images.remove(img)
                break
    # 2. 插入新图（等比缩放）
    for label, (row, col) in anchors.items():
        if label not in unit_images:
            continue
        nimg = XLImage(io.BytesIO(unit_images[label]))
        max_w, max_h = DISPIMG_DEFAULTS[label][1], DISPIMG_DEFAULTS[label][2]
        iw, ih = nimg.width, nimg.height
        if iw and ih:
            scale = min(max_w / iw, max_h / ih, 1.0)
            if scale < 1.0:
                nimg.width = max(int(iw * scale), 1)
                nimg.height = max(int(ih * scale), 1)
        ws.add_image(nimg, f"{get_column_letter(col)}{row}")


def _purge_data_area(ws):
    """兜底清空 R6~R23 数据区所有单元格的值和样式（防止模板 SUMIFS 等公式污染 → #VALUE!）"""
    from openpyxl.styles import Border as _Border
    from openpyxl.cell.cell import MergedCell
    empty = _Border()
    # 1. 先解除所有包含 R6~R23 的合并（其他区域的合并保持；用 try 跳过 delete_cols 后的"幽灵"合并）
    safe_merged = []
    for mr in list(ws.merged_cells.ranges):
        # 只关心跟 R6~R23 有关联的合并
        if mr.min_row <= 23 and mr.max_row >= 6:
            try:
                ws.unmerge_cells(str(mr))
            except Exception:
                pass  # 忽略 delete_cols 后的"幽灵"合并区域
        else:
            safe_merged.append(str(mr))
    # 2. 清空 R6~R23 所有单元格值（跳过 MergedCell）
    for r in range(6, 24):
        for c in range(1, 13):
            cell = ws.cell(row=r, column=c)
            if isinstance(cell, MergedCell):
                continue
            cell.value = None
            cell.border = empty
    # 3. 恢复 R6:R7 合并（让后续尺码表头能正常合并）
    for col in range(1, 13):
        coord = f"{get_column_letter(col)}6:{get_column_letter(col)}7"
        if coord not in {str(r) for r in ws.merged_cells.ranges}:
            try:
                ws.merge_cells(coord)
            except Exception:
                pass


# ---------- 订单 sheet ----------

def _build_order_sheet(wb, unit: OrderUnit):
    ws = wb["订单"]

    # 0. 只保留 A~L 列（删除 THD 模板 M 列后的 裁床数量/汇总 区）
    if ws.max_column > 13:
        ws.delete_cols(14, ws.max_column - 13)

    # 0.5 兜底：清空整个数据区所有公式/旧值（防止模板污染 → #VALUE! / #REF!）
    #    数据区 = R6(表头) ~ R23(原"总计"行)
    _purge_data_area(ws)

    # 2. 先插列（尺码列表头 C~H 固定 6 列，>6 个尺码需插列）
    #    ⚠️ 必须先插列再写头部：insert_cols 会把 H 列及其后的列右移，
    #       若先写 H2/H3/H5 会被挤到 K 列（导致名称/品牌/日期丢失）
    sizes = unit.sizes
    need_insert = len(sizes) - 6
    if need_insert > 0:
        ws.insert_cols(8, need_insert)   # H 列前插

    # 1. 头部字段
    ws["A1"] = "生产订单"
    ws["B2"] = unit.style          # 款号
    ws["H2"] = unit.product_name   # 名称
    ws["B3"] = unit.order_no       # 订单号
    ws["H3"] = unit.config_brand   # 品牌
    ws["B4"] = unit.pattern        # 做货纸样编号
    ws["G4"] = ""                  # 去掉「店铺」标签
    ws["H4"] = ""
    ws["B5"] = unit.factory        # 加工厂
    ws["H5"] = unit.date_str       # 日期

    # 2.5 比例列加宽到原来的 3 倍（用户要求：避免比例字符串被截断）
    ratio_col = 3 + len(sizes)   # 比例列索引
    ratio_col_letter = get_column_letter(ratio_col)
    original_width = ws.column_dimensions[ratio_col_letter].width or 13.0
    ws.column_dimensions[ratio_col_letter].width = original_width * 3
    # 重写尺码表头（R6 为合并左上角，R7 是 MergedCell 不可写）
    for i, s in enumerate(sizes):
        col = 3 + i
        ws.cell(row=6, column=col).value = s
    # 新插入的尺码列补合并 R6:R7 保持表头样式
    if need_insert > 0:
        for i in range(6, 6 + need_insert):
            col = 3 + i
            if not any(str(r) == f"{get_column_letter(col)}6:{get_column_letter(col)}7"
                       for r in ws.merged_cells.ranges):
                ws.merge_cells(start_row=6, start_column=col, end_row=7, end_column=col)
    # 清理多余旧表头（仅左上角）
    for i in range(len(sizes), 6 + need_insert):
        col = 3 + i
        ws.cell(row=6, column=col).value = None

    # 3. 数据行（R8 起）+ 汇总行（隔一行）
    start = 8
    n = len(unit.color_rows)
    # 确保有足够行：THD 模板固定内容从 R24 起，数据区 R8~R23 = 16 行
    fixed_top = 24
    avail = fixed_top - start
    need_rows = n + 2   # 数据行 + 空行 + 汇总行
    if need_rows > avail:
        ws.insert_rows(start + avail, need_rows - avail)   # 在固定内容前插行
        _shift_images_down(ws, start + avail, need_rows - avail)  # 图片随固定内容下移
    for idx, cr in enumerate(unit.color_rows):
        r = start + idx
        ws.cell(row=r, column=1, value=cr.color)
        ws.cell(row=r, column=2, value=cr.bag_spec)
        for i, s in enumerate(sizes):
            ws.cell(row=r, column=3 + i, value=cr.qty_by_size.get(s) or None)
        ws.cell(row=r, column=3 + len(sizes), value=cr.ratio)      # 比例
        ws.cell(row=r, column=4 + len(sizes), value=cr.total)      # 总计
        ws.cell(row=r, column=5 + len(sizes), value=cr.rolls)      # 布料条数
        ws.cell(row=r, column=6 + len(sizes), value=cr.remark)     # 备注
        # 边框
        for c in range(1, 7 + len(sizes)):
            cell = ws.cell(row=r, column=c)
            border = cell.border
            if border is None or border.left is None or not border.left.style:
                cell.border = BOX

    # 汇总行（数据行后隔一行）：总件数 + 布料总条数
    # 先清空空行（数据行与汇总行之间的行可能残留模板公式 → 防止 #REF!)
    blank_row = start + n
    for c in range(1, 7 + len(sizes)):
        cell = ws.cell(row=blank_row, column=c)
        cell.value = None
        cell.border = Border()
    total_row = start + n + 1
    ws.cell(row=total_row, column=1, value="总计")
    ws.cell(row=total_row, column=2, value="")   # 包装袋规格留空
    # 不再写各尺码合计（按用户要求：只显示总件数 + 布料总条数）
    ws.cell(row=total_row, column=3 + len(sizes), value="")                     # 比例留空
    ws.cell(row=total_row, column=4 + len(sizes), value=unit.total_qty)         # 总件数
    ws.cell(row=total_row, column=5 + len(sizes), value=sum(cr.rolls for cr in unit.color_rows))  # 布料总条数
    ws.cell(row=total_row, column=6 + len(sizes), value="")                     # 备注留空
    for c in range(1, 7 + len(sizes)):
        cell = ws.cell(row=total_row, column=c)
        cell.border = BOX
    # 字体不加粗，与颜色行一致

    # 4. 洗水唛内容（THD 模板 R36 标签 / R37 内容）
    for r in range(30, min(ws.max_row, 45) + 1):
        if ws.cell(row=r, column=1).value and "洗水唛" in str(ws.cell(row=r, column=1).value):
            ws.cell(row=r + 1, column=1, value=unit.wash_label)
            break

    # 5. 清空数据区残留的空白行（原模板 R8~R23 可能含公式/旧值 → 防止 #VALUE）
    last_used_row = total_row
    for r in range(last_used_row + 1, fixed_top):
        for c in range(1, 7 + len(sizes)):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.border = Border()

    # 6. 仅当输入文件/网页提供了替换图时，替换模板中对应位置的图片（模板默认图保留）
    _replace_cell_images(ws, getattr(unit, "images", None) or {})
    return ws


# ---------- 布料申购单 sheet ----------

def _build_supplier_sheet(wb, unit: OrderUnit, group):
    """从 PPY 模板布料申购单复制样式并生成一个布行申购单 sheet"""
    src = openpyxl.load_workbook(PPY_TEMPLATE, data_only=False)
    src_ws = src["布料申购单"]
    new_ws = wb.create_sheet(title=f"布料申购单-{group.supplier}")

    # 复制列宽
    for col, dim in src_ws.column_dimensions.items():
        new_ws.column_dimensions[col].width = dim.width
    # 复制行高 + 单元格值/样式（仅左块 A~D 列 + 表头，右块 H~K 不复制）
    for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row, max_col=4):
        for c in row:
            if c.value is not None:
                new_ws.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                new_ws.cell(row=c.row, column=c.column).font = copy(c.font)
                new_ws.cell(row=c.row, column=c.column).border = copy(c.border)
                new_ws.cell(row=c.row, column=c.column).fill = copy(c.fill)
                new_ws.cell(row=c.row, column=c.column).alignment = copy(c.alignment)
            new_ws.row_dimensions[c.row].height = src_ws.row_dimensions[c.row].height
    # 复制左块合并单元格（A1:D1, A2:D2... 检查）
    for mr in src_ws.merged_cells.ranges:
        if mr.min_col <= 4:   # 只复制左块
            new_ws.merge_cells(str(mr))

    # 填头部
    new_ws["B2"] = unit.order_no
    new_ws["D2"] = unit.style
    new_ws["B3"] = ""            # 货号 留空
    new_ws["D3"] = ""            # 重量 留空
    new_ws["B4"] = unit.fabric   # 品名
    new_ws["D4"] = unit.factory  # 代加工工厂
    new_ws["B8"] = unit.config_applicant
    new_ws["D8"] = unit.date_str

    # 数据行（R6 起，模板 R6 是第 1 条数据）
    data_start = 6
    rows = group.rows
    need = len(rows)
    # 模板数据行从 R6 到 R6（1 行），合计在 R7，申请人在 R8
    # 插入额外行使数据区容纳所有颜色
    if need > 1:
        new_ws.insert_rows(data_start + 1, need - 1)
    # 以 R6 为样式基准，复制到所有数据行（修复 insert_rows 后样式不一致）
    for i, sr in enumerate(rows):
        r = data_start + i
        new_ws.cell(row=r, column=1, value=sr.color)
        new_ws.cell(row=r, column=2, value=sr.rolls)
        new_ws.cell(row=r, column=3, value=sr.supplier)
        new_ws.cell(row=r, column=4, value=unit.config_account_code)
        # 复制 R6 的样式（边框/字体/对齐/填充）到 R(i)，确保与 R6 一致
        for c in range(1, 5):
            src_c = new_ws.cell(row=data_start, column=c)
            dst_c = new_ws.cell(row=r, column=c)
            if src_c.has_style:
                dst_c.font = copy(src_c.font)
                dst_c.alignment = copy(src_c.alignment)
                dst_c.fill = copy(src_c.fill)
                dst_c.border = copy(src_c.border)
        new_ws.row_dimensions[r].height = new_ws.row_dimensions[data_start].height
    # 合计行
    total_row = data_start + need
    new_ws.cell(row=total_row, column=1, value="合计")
    new_ws.cell(row=total_row, column=2, value=group.total_rolls)
    # 申请人行（若插行后申请人行被下移，重定位）
    new_ws.cell(row=total_row + 1, column=1, value="申请人")
    new_ws.cell(row=total_row + 1, column=2, value=unit.config_applicant)
    new_ws.cell(row=total_row + 1, column=3, value="申请日期")
    new_ws.cell(row=total_row + 1, column=4, value=unit.date_str)
    # 合计行 / 申请人行 应用与数据行一致的样式（字体/对齐/边框/行高）
    for rr in (total_row, total_row + 1):
        for c in range(1, 5):
            src_c = new_ws.cell(row=data_start, column=c)
            dst_c = new_ws.cell(row=rr, column=c)
            if src_c.has_style:
                dst_c.font = copy(src_c.font)
                dst_c.alignment = copy(src_c.alignment)
                dst_c.fill = copy(src_c.fill)
                dst_c.border = copy(src_c.border)
        new_ws.row_dimensions[rr].height = new_ws.row_dimensions[data_start].height
    return new_ws


# ---------- 汇总明细表 ----------

def build_summary(split_result: SplitResult, date_str: str, out_dir: str) -> str:
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "文件汇总"
    headers1 = ["订单号", "款号", "产品名称", "类型", "工厂", "涉及布行", "颜色数", "总件数", "布料总条数", "生成文件名"]
    ws1.append(headers1)
    for u in split_result.units:
        suppliers = "、".join(g.supplier for g in u.supplier_groups)
        ws1.append([u.order_no, u.style, u.product_name, u.type_str, u.factory,
                    suppliers, len(u.color_rows), u.total_qty,
                    sum(g.total_rolls for g in u.supplier_groups), u.file_name])

    ws2 = wb.create_sheet("布行明细")
    headers2 = ["订单号", "布行", "品名", "颜色数", "条数合计"]
    ws2.append(headers2)
    for u in split_result.units:
        for g in u.supplier_groups:
            ws2.append([u.order_no, g.supplier, g.fabric, len(g.rows), g.total_rolls])

    for ws, ncol in ((ws1, len(headers1)), (ws2, len(headers2))):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = Font(bold=True)
            cell.border = BOX
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = BOX
        ws.freeze_panes = "A2"

    path = os.path.join(out_dir, f"{date_str}_采购单汇总.xlsx")
    wb.save(path)
    return path


# ---------- 入口 ----------

def generate_all(data: InputData, split_result: SplitResult, out_dir: str) -> List[str]:
    os.makedirs(out_dir, exist_ok=True)
    date_str = split_result.units[0].order_no[3:11] if split_result.units else data.input_filename
    paths = []
    for unit in split_result.units:
        unit.config_brand = data.config.brand
        unit.config_applicant = data.config.applicant
        unit.config_account_code = data.config.account_code
        unit.config_office = data.config.office
        unit.date_str = date_str
        unit.images = data.images   # 输入文件「图片」sheet 的图（可能为空 → 用默认）
        # 复制 THD 模板
        wb = openpyxl.load_workbook(THD_TEMPLATE, data_only=False)
        # 删除除「订单」外所有 sheet
        for ws in list(wb.worksheets):
            if ws.title != "订单":
                wb.remove(ws)
        _build_order_sheet(wb, unit)
        for g in unit.supplier_groups:
            _build_supplier_sheet(wb, unit, g)
        path = os.path.join(out_dir, unit.file_name)
        wb.save(path)
        paths.append(path)

    # 汇总表
    build_summary(split_result, date_str, out_dir)
    return paths


def make_zip(out_dir: str, zip_path: str, date_str: str) -> str:
    """打包 out_dir 下所有 xlsx 为 zip"""
    files = [f for f in os.listdir(out_dir) if f.endswith(".xlsx")]
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(os.path.join(out_dir, f), arcname=f)
    return zip_path
