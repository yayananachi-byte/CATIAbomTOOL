"""将 CATIA 图片清单 BOM 写入“基础数据管制表”的产品区（A:AN）。

依赖：openpyxl、Pillow（均随 Codex 运行环境提供）。

示例：
    # 不带参数运行：依次选择 BOM、模板和输出文件
    python bom_to_control_table.py

    # 命令行指定文件：适合批量处理
    python bom_to_control_table.py \
      --bom CATIA_图片清单_STU-0000509373_03_20260901_135340.xlsx \
      --template 客户XXX-项目名XXX-基础数据管制表-日期.xlsx

脚本只会写入产品区 A:AN；AO:DH 的模具、检具、夹具、成本等内容保持模板原样。
当 BOM 记录数超过模板现有的 16 个记录槽位时，脚本会在汇总区前按每项 6 行扩展记录区。
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any, Callable, Optional, Sequence

import openpyxl
from PIL import Image, ImageDraw
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.cell_range import CellRange


DATA_START_ROW = 10
ROWS_PER_RECORD = 6
PRODUCT_LAST_COL = 40  # AN
TEMPLATE_IMAGE_CM = 8.57
TEMPLATE_IMAGE_PX = round(TEMPLATE_IMAGE_CM * 96 / 2.54)
IMAGE_MARGIN = 0.10
HIERARCHY_COLUMNS = ("D", "E", "F", "G", "H", "I", "J")

# 汽车常见标准件。名称命中这些词（中英文均可）或零件号以 Q/X 开头时，
# 将重量填写到“标准件净重/kg”(W)，否则填写到“产品净重/kg”(X)。
STANDARD_PART_KEYWORDS = (
    "标准件", "螺栓", "螺母", "螺钉", "螺丝", "垫圈", "垫片", "铆钉",
    "销", "卡扣", "夹子", "弹簧", "轴承", "紧固件", "焊钉", "螺柱",
    "bolt", "nut", "screw", "washer", "rivet", "pin", "clip", "clamp",
    "fastener", "stud", "bearing", "spring", "snap", "cotter",
)


class ConversionCancelled(RuntimeError):
    """用户在安全写入边界请求取消模板转换。"""


def _check_conversion_cancel(cancel_requested: Optional[Callable[[], bool]]) -> None:
    if cancel_requested is not None and cancel_requested():
        raise ConversionCancelled("已收到取消请求，尚未生成输出文件。")


def _emit_progress(
    progress_callback: Optional[Callable[[int, str], None]],
    percent: int,
    message: str,
) -> None:
    """向 GUI 报告进度；界面更新失败不影响转换。"""
    if progress_callback is None:
        return
    try:
        progress_callback(max(0, min(100, int(percent))), message)
    except Exception:
        pass


def clean_text(value: Any) -> str | None:
    """返回清理后的文本；空、None 和字符串 None 统一为空。"""
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.lower() == "none":
        return None
    return text


def trim_part_name(value: Any) -> str | None:
    return clean_text(value)


def extract_thickness(value: Any) -> float | None:
    """从“0.700 mm”一类文本中取出厚度；未识别时留空。"""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = clean_text(value)
    if not text:
        return None
    match = re.search(r"(?<![A-Za-z])(-?\d+(?:\.\d+)?)\s*(?:mm)?\b", text, flags=re.I)
    return float(match.group(1)) if match else None


def get_value(record: dict[str, Any], header: str) -> Any:
    value = record.get(header)
    if isinstance(value, str):
        return clean_text(value)
    return value


def as_number(value: Any) -> float | None:
    """安全转换数值；用于质量字段统一保留两位小数。"""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def read_bom(bom_path: Path) -> tuple[list[dict[str, Any]], dict[int, ExcelImage]]:
    """读取 CATIA 图片清单，并按原树顺序返回记录与行号对应的缩略图。"""
    workbook = openpyxl.load_workbook(bom_path)
    sheet = workbook.active

    headers = {
        clean_text(cell.value): cell.column
        for cell in sheet[2]
        if clean_text(cell.value)
    }
    required_headers = {"树编号", "层级", "类型", "零件号", "版本", "零件名称", "实例名称"}
    missing = required_headers - set(headers)
    if missing:
        raise ValueError(f"BOM 缺少必需列：{'、'.join(sorted(missing))}")

    # 旧版简版 CATIA_BOM 只带结构、图片和 XYZ，无法可靠填写基础信息。
    # 请用更新后的 catia_bom_export_v2.py 或 catia_bom_screenshot.py 重新导出。
    physical_headers = {
        "单件质量(kg)", "包装包络长(mm)", "包装包络宽(mm)", "包装包络高(mm)",
        "名义厚度", "本级数量",
    }
    missing_physical = physical_headers - set(headers)
    if missing_physical:
        raise ValueError(
            "所选文件是简版 CATIA_BOM，缺少基础信息列："
            f"{'、'.join(sorted(missing_physical))}。请先运行更新后的 catia_bom_export_v2.py "
            "或 catia_bom_screenshot.py 重新导出，再选择新生成的 Excel。"
        )

    records: list[dict[str, Any]] = []
    for row_index in range(3, sheet.max_row + 1):
        tree_no = clean_text(sheet.cell(row_index, headers["树编号"]).value)
        if not tree_no:
            continue
        record = {header: sheet.cell(row_index, col).value for header, col in headers.items()}
        record["_source_row"] = row_index
        record["_tree_no"] = tree_no
        record["_path"] = tree_no.split(".")
        records.append(record)

    if not records:
        raise ValueError("BOM 中未找到任何产品结构记录。")

    images_by_row: dict[int, ExcelImage] = {}
    for image in sheet._images:
        anchor = image.anchor
        if hasattr(anchor, "_from"):
            images_by_row.setdefault(anchor._from.row + 1, image)

    return records, images_by_row


def locate_summary_start(sheet) -> int:
    """从首个数据块后的汇总标签定位汇总区；模板默认是第 106 行。"""
    for row in range(DATA_START_ROW + ROWS_PER_RECORD, sheet.max_row + 1):
        value = clean_text(sheet.cell(row, 4).value)
        if value in {"总成开发数/个：", "总成开发数/个:"}:
            return row
    raise ValueError("未找到模板汇总区起始行（“总成开发数/个”）。")


def shift_merges_for_insert(sheet, insert_row: int, amount: int) -> None:
    """openpyxl 插行不会移动合并区域，需同步移动插入点以下的合并范围。"""
    original_merges = [str(merged) for merged in sheet.merged_cells.ranges]
    sheet.insert_rows(insert_row, amount)

    moved_merges: list[CellRange] = []
    for merge_ref in original_merges:
        merged = CellRange(merge_ref)
        if merged.min_row >= insert_row:
            merged.shift(row_shift=amount, col_shift=0)
        moved_merges.append(merged)
    sheet.merged_cells.ranges = set(moved_merges)


def copy_record_block(sheet, source_start: int, target_start: int) -> None:
    """复制一个 6 行数据槽位的样式、公式和纵向合并结构。"""
    source_merges = [
        str(merged)
        for merged in sheet.merged_cells.ranges
        if merged.min_row >= source_start and merged.max_row < source_start + ROWS_PER_RECORD
    ]

    for row_offset in range(ROWS_PER_RECORD):
        source_row = source_start + row_offset
        target_row = target_start + row_offset
        if source_row in sheet.row_dimensions:
            sheet.row_dimensions[target_row] = copy(sheet.row_dimensions[source_row])

        for col in range(1, sheet.max_column + 1):
            source_cell = sheet.cell(source_row, col)
            if isinstance(source_cell, MergedCell):
                continue
            target_cell = sheet.cell(target_row, col)
            if source_cell.has_style:
                target_cell._style = copy(source_cell._style)
            if source_cell.number_format:
                target_cell.number_format = source_cell.number_format
            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.protection = copy(source_cell.protection)

            if source_cell.data_type == "f":
                target_cell.value = Translator(
                    source_cell.value, origin=source_cell.coordinate
                ).translate_formula(target_cell.coordinate)
            else:
                target_cell.value = source_cell.value

    for merge_ref in source_merges:
        merged = CellRange(merge_ref)
        merged.shift(row_shift=target_start - source_start, col_shift=0)
        sheet.merge_cells(str(merged))


def ensure_record_capacity(sheet, record_count: int) -> tuple[list[int], int]:
    """按 BOM 记录数确保模板有足够的 6 行记录槽位。"""
    summary_start = locate_summary_start(sheet)
    existing_capacity = (summary_start - DATA_START_ROW) // ROWS_PER_RECORD
    if record_count > existing_capacity:
        extra_records = record_count - existing_capacity
        extra_rows = extra_records * ROWS_PER_RECORD
        shift_merges_for_insert(sheet, summary_start, extra_rows)

        exemplar_start = DATA_START_ROW + (existing_capacity - 1) * ROWS_PER_RECORD
        for index in range(extra_records):
            target_start = summary_start + index * ROWS_PER_RECORD
            copy_record_block(sheet, exemplar_start, target_start)

        # 插行后 openpyxl 不会更新公式引用。只平移原汇总区及其下方
        # 指向“旧汇总区”的单元格引用；产品数据区（10:105）引用保持不变。
        new_summary_start = summary_start + extra_rows
        shift_summary_references(
            sheet,
            first_summary_row=summary_start,
            new_summary_start=new_summary_start,
            row_offset=extra_rows,
        )
        summary_start = new_summary_start

        # 模具类别和出件数的模板下拉范围同步扩展，避免新增记录没有校验。
        last_data_row = summary_start - 1
        for validation in sheet.data_validations.dataValidation:
            validation.sqref = re.sub(
                r"(\$?[A-Z]{1,3}\$?10:\$?[A-Z]{1,3}\$?)105\b",
                rf"\g<1>{last_data_row}",
                str(validation.sqref),
            )

    slot_starts = list(range(DATA_START_ROW, summary_start, ROWS_PER_RECORD))
    return slot_starts, summary_start


def clear_product_area(sheet, slot_starts: list[int]) -> None:
    """清空 A:AN 的模板占位值，不触及 AO:DH。"""
    for start_row in slot_starts:
        for row in range(start_row, start_row + ROWS_PER_RECORD):
            for col in range(1, PRODUCT_LAST_COL + 1):
                cell = sheet.cell(row, col)
                if not isinstance(cell, MergedCell):
                    cell.value = None


def is_standard_part(record: dict[str, Any]) -> bool:
    """按照零件名称和常见标准件号规则识别汽车标准件。"""
    part_no = clean_text(record.get("零件号")) or ""
    name = (clean_text(record.get("零件名称")) or "").casefold()
    return bool(re.match(r"^[QX]\d", part_no, flags=re.I)) or any(
        keyword.casefold() in name for keyword in STANDARD_PART_KEYWORDS
    )


def is_root_assembly(record: dict[str, Any]) -> bool:
    return clean_text(record.get("类型")) == "根产品"


def is_subassembly(record: dict[str, Any]) -> bool:
    return clean_text(record.get("类型")) == "子总成"


def is_part(record: dict[str, Any]) -> bool:
    return clean_text(record.get("类型")) == "零件"


def descendant_last_index(records: list[dict[str, Any]], index: int) -> int:
    """返回当前节点子树的最后一条记录（BOM 保持深度优先顺序）。"""
    path = records[index]["_path"]
    for candidate in range(index + 1, len(records)):
        if records[candidate]["_path"][: len(path)] != path:
            return candidate - 1
    return len(records) - 1


def set_value(sheet, row: int, column: str, value: Any) -> None:
    if value is not None and value != "":
        sheet[f"{column}{row}"].value = value


def set_mass_value(sheet, row: int, column: str, value: Any) -> None:
    """写入 kg 质量，并统一以两位小数显示。"""
    mass = as_number(value)
    if mass is None:
        return
    cell = sheet[f"{column}{row}"]
    cell.value = round(mass, 2)
    cell.number_format = "0.00"


def crop_to_square(image_bytes: bytes) -> BytesIO:
    """按 CATIA 的近似纯色背景定位全部可见零件，再补成正方形。"""
    with Image.open(BytesIO(image_bytes)) as source:
        original = source.convert("RGB")
    width, height = original.size
    preview = original.copy()
    preview.thumbnail((800, 800))
    pw, ph = preview.size
    pixels = preview.load()
    samples = [
        pixels[int(pw * x), int(ph * y)]
        for x, y in ((0.04, 0.5), (0.96, 0.5), (0.5, 0.04), (0.5, 0.96))
    ]
    background = tuple(sorted(sample[channel] for sample in samples)[2] for channel in range(3))
    consistent = sum(
        max(abs(color[c] - background[c]) for c in range(3)) < 35
        for color in samples
    ) >= 3
    bounds = None
    has_tree = False
    if consistent:
        # 旧版 CATIA 截图可能带左侧产品树：左边缘多行持续有图标/文字。
        tree_rows = sum(
            any(max(abs(pixels[x, y][c] - background[c]) for c in range(3)) > 45
                for x in range(0, max(1, int(pw * 0.06)), 3))
            for y in range(0, ph, 3)
        )
        has_tree = 0.12 * (ph / 3) < tree_rows < 0.80 * (ph / 3)
        xmin, ymin, xmax, ymax = pw, ph, -1, -1
        background_pixels = 0
        for y in range(ph):
            for x in range(pw):
                color = pixels[x, y]
                difference = max(abs(color[c] - background[c]) for c in range(3))
                if difference <= 45:
                    background_pixels += 1
                    continue
                if has_tree and x < 0.30 * pw:
                    continue
                # CATIA 的方向罗盘与坐标轴位于视窗右侧的两个角落。
                if (x >= 0.84 * pw and y <= 0.18 * ph) or (x >= 0.90 * pw and y >= 0.86 * ph):
                    continue
                xmin, ymin = min(xmin, x), min(ymin, y)
                xmax, ymax = max(xmax, x), max(ymax, y)
        if (background_pixels >= 0.45 * pw * ph and
                xmax >= xmin and ymax >= ymin and
                xmax - xmin >= 0.02 * pw and ymax - ymin >= 0.02 * ph):
            bounds = (
                math.floor(xmin * width / pw), math.floor(ymin * height / ph),
                math.ceil((xmax + 1) * width / pw), math.ceil((ymax + 1) * height / ph),
            )

    if bounds is None:
        bounds = (0, 0, width, height)
    if has_tree:
        # 裁切框可能为补足正方形而向左扩展，需同时抹掉残余产品树。
        ImageDraw.Draw(original).rectangle((0, 0, int(width * 0.30), height), fill=background)
    left, top, right, bottom = bounds
    box_width, box_height = right - left, bottom - top
    side = max(1, math.ceil(max(box_width, box_height) * (1 + 2 * IMAGE_MARGIN)))
    center_x, center_y = (left + right) / 2, (top + bottom) / 2
    square_left = math.floor(center_x - side / 2)
    square_top = math.floor(center_y - side / 2)
    square = Image.new("RGB", (side, side), background)
    crop_box = (max(0, square_left), max(0, square_top),
                min(width, square_left + side), min(height, square_top + side))
    if crop_box[2] > crop_box[0] and crop_box[3] > crop_box[1]:
        square.paste(original.crop(crop_box),
                     (crop_box[0] - square_left, crop_box[1] - square_top))
    result = BytesIO()
    square.save(result, format="PNG")
    result.seek(0)
    return result


def copy_thumbnail(sheet, source_image: ExcelImage | None, target_cell: str) -> None:
    if source_image is None:
        return
    target_image = ExcelImage(crop_to_square(source_image._data()))
    target_image.width = TEMPLATE_IMAGE_PX
    target_image.height = TEMPLATE_IMAGE_PX
    # AN 为 6 行合并单元格；让 8.57 cm 的图片完整落在该单元格内。
    sheet.column_dimensions["AN"].width = max(sheet.column_dimensions["AN"].width or 0, 47)
    start_row = sheet[target_cell].row
    for row in range(start_row, start_row + ROWS_PER_RECORD):
        sheet.row_dimensions[row].height = max(sheet.row_dimensions[row].height or 0, 41)
    sheet.add_image(target_image, target_cell)


def write_record(
    sheet,
    target_row: int,
    record: dict[str, Any],
    source_image: ExcelImage | None,
    geely_mode: bool = False,
) -> None:
    path = record["_path"]
    depth = len(path) - 1
    if depth > 6:
        raise ValueError(
            f"树编号 {record['_tree_no']} 的层级为 {depth}，超过模板支持的 6 层（E:J）。"
        )

    part_no = clean_text(record.get("零件号"))
    part_name = trim_part_name(record.get("零件名称"))
    # 吉利原表的名称在导出时已去掉 / 后的版本；转换旧表时也做一次兼容。
    identifier = (part_name or "").partition("/")[0].strip() if geely_mode else part_no
    material = clean_text(record.get("PartBody Material")) or clean_text(record.get("属性 Material"))
    thickness = extract_thickness(record.get("名义厚度"))
    piece_mass = get_value(record, "单件质量(kg)")

    # 每条记录只写入所属的那一级：根总成在 D，一级子项在 E，二级子项在 F，依此类推。
    # 例如 1 / 1.1 / 1.1.1 分别写在 D / E / F，不在同一行重复祖先层级。
    set_value(sheet, target_row, HIERARCHY_COLUMNS[depth], record["_tree_no"])

    # 总成、分总成、零件号三列互斥，避免在同一条记录重复填写编号。
    if is_root_assembly(record):
        set_value(sheet, target_row, "M", identifier)
    elif is_subassembly(record):
        set_value(sheet, target_row, "N", identifier)
    elif is_part(record):
        set_value(sheet, target_row, "O", identifier)

    if not geely_mode:
        set_value(sheet, target_row, "Q", part_name)
    if is_part(record):
        # 完整图片清单中的“本级数量”就是零件单耗；缺失时按 1 件处理。
        set_value(sheet, target_row, "R", get_value(record, "本级数量") or 1)
        if is_standard_part(record):
            set_mass_value(sheet, target_row, "W", piece_mass)
        else:
            set_mass_value(sheet, target_row, "X", piece_mass)

        set_value(sheet, target_row, "Y", material)
        set_value(sheet, target_row, "Z", thickness)
        set_value(sheet, target_row, "AB", get_value(record, "包装包络长(mm)"))
        set_value(sheet, target_row, "AC", get_value(record, "包装包络宽(mm)"))
        set_value(sheet, target_row, "AD", get_value(record, "包装包络高(mm)"))
    set_value(sheet, target_row, "AL", clean_text(record.get("版本")))
    copy_thumbnail(sheet, source_image, f"AN{target_row}")


def write_assembly_mass_formulas(
    sheet, records: list[dict[str, Any]], target_slots: list[int]
) -> None:
    """用 Excel 公式统计总成和分总成的净重。

    W、X 是每件标准件/产品净重，R 是单耗。公式只逐项引用实际零件所在的
    记录行，不引用合并单元格内的空白行；因此不会多算，也不会重复计算分总成。
    """
    for index, record in enumerate(records):
        if not (is_root_assembly(record) or is_subassembly(record)):
            continue
        last_index = descendant_last_index(records, index)
        target_row = target_slots[index]
        terms: list[str] = []
        for descendant_index in range(index + 1, last_index + 1):
            descendant = records[descendant_index]
            if not is_part(descendant):
                continue
            mass_column = "W" if is_standard_part(descendant) else "X"
            part_row = target_slots[descendant_index]
            terms.append(f"SUMPRODUCT($R${part_row},${mass_column}${part_row})")
        formula = f"=ROUND({'+'.join(terms)},2)" if terms else "=0"
        target_cell = sheet[f"{'U' if is_root_assembly(record) else 'V'}{target_row}"]
        target_cell.value = formula
        target_cell.number_format = "0.00"


def write_summary_formulas(
    sheet, summary_start: int, records: list[dict[str, Any]], target_slots: list[int]
) -> None:
    """填写汇总区的总成开发数与总成净重。"""
    if not target_slots:
        return
    last_row = target_slots[-1]
    # 每条记录块的 D 单元格纵向合并，不能用 Excel 直接计数；由 BOM 根节点数准确填写。
    sheet[f"L{summary_start}"] = sum(1 for record in records if is_root_assembly(record))
    # U 仅在根总成行有公式，因此可直接求和，适用于一个或多个根总成。
    summary_mass = sheet[f"Q{summary_start}"]
    summary_mass.value = f"=ROUND(SUM($U$10:$U${last_row}),2)"
    summary_mass.number_format = "0.00"


def update_summary_formula_ranges(sheet, summary_start: int) -> None:
    """扩容后，将模板汇总公式的原始数据行终点 105 更新为新终点。"""
    last_data_row = summary_start - 1
    pattern = re.compile(r"(?P<prefix>[A-Z]{1,3}10:[A-Z]{1,3})105\b")
    for row in sheet.iter_rows(min_row=summary_start, max_row=sheet.max_row):
        for cell in row:
            if cell.data_type == "f" and isinstance(cell.value, str):
                cell.value = pattern.sub(rf"\g<prefix>{last_data_row}", cell.value)


def shift_summary_references(
    sheet, first_summary_row: int, new_summary_start: int, row_offset: int
) -> None:
    """平移插行后位于汇总区的公式中、指向原汇总区的引用。"""
    cell_ref = re.compile(r"(?P<col>\$?[A-Z]{1,3}\$?)(?P<row>\d+)\b")

    def replace_reference(match: re.Match[str]) -> str:
        row = int(match.group("row"))
        if row >= first_summary_row:
            return f"{match.group('col')}{row + row_offset}"
        return match.group(0)

    for row in sheet.iter_rows(min_row=new_summary_start, max_row=sheet.max_row):
        for cell in row:
            if cell.data_type == "f" and isinstance(cell.value, str):
                cell.value = cell_ref.sub(replace_reference, cell.value)


def rebase_root_numbers(
    records: list[dict[str, Any]], next_root_number: int
) -> tuple[list[dict[str, Any]], int]:
    """为合并后的 BOM 重新编号根节点，保持每棵树相互独立。

    单个 CATIA BOM 通常都从 ``1`` 开始。如果直接把多份 BOM 拼接，第一份
    BOM 的根 ``1`` 会把后续 BOM 的根 ``1`` 当成自己的后代，进而导致总成
    重量公式错误。这里只替换路径的第一级编号，子级编号与原 BOM 保持不变。
    """
    root_number_map: dict[str, str] = {}
    rebased_records: list[dict[str, Any]] = []

    for record in records:
        original_path = list(record["_path"])
        if not original_path:
            raise ValueError("BOM 中存在空的树编号，无法合并。")
        original_root = original_path[0]
        if original_root not in root_number_map:
            root_number_map[original_root] = str(next_root_number)
            next_root_number += 1
        rebased_path = [root_number_map[original_root], *original_path[1:]]
        rebased = dict(record)
        rebased_tree_no = ".".join(rebased_path)
        rebased["_path"] = rebased_path
        rebased["_tree_no"] = rebased_tree_no
        # 写入产品区时会用到“树编号”字段，需与内部路径保持一致。
        rebased["树编号"] = rebased_tree_no
        rebased_records.append(rebased)

    return rebased_records, next_root_number


def read_and_merge_boms(
    bom_paths: Sequence[Path],
    progress_callback: Optional[Callable[[int, str], None]] = None,
    cancel_requested: Optional[Callable[[], bool]] = None,
) -> list[dict[str, Any]]:
    """读取多份 CATIA 图片清单并合并成一棵安全的记录序列。"""
    if not bom_paths:
        raise ValueError("请至少选择一份 CATIA BOM 图片清单。")

    combined_records: list[dict[str, Any]] = []
    next_root_number = 1
    total = len(bom_paths)
    for index, bom_path in enumerate(bom_paths, 1):
        _check_conversion_cancel(cancel_requested)
        _emit_progress(
            progress_callback,
            round((index - 1) * 35 / total),
            f"正在读取 BOM {index}/{total}：{bom_path.name}",
        )
        records, images_by_row = read_bom(bom_path)
        rebased_records, next_root_number = rebase_root_numbers(records, next_root_number)
        for record in rebased_records:
            record["_source_image"] = images_by_row.get(record["_source_row"])
        combined_records.extend(rebased_records)

    if not combined_records:
        raise ValueError("所选 BOM 中没有可转换的产品结构记录。")
    return combined_records


def convert_many(
    bom_paths: Sequence[Path],
    template_path: Path,
    output_path: Path,
    progress_callback: Optional[Callable[[int, str], None]] = None,
    cancel_requested: Optional[Callable[[], bool]] = None,
    geely_mode: bool = False,
) -> None:
    """将一份或多份 CATIA BOM 合并写入同一份基础数据管制表模板。"""
    normalized_bom_paths = [Path(path).resolve() for path in bom_paths]
    normalized_template_path = Path(template_path).resolve()
    normalized_output_path = Path(output_path).resolve()
    if not normalized_bom_paths:
        raise ValueError("请至少选择一份 CATIA BOM 图片清单。")
    missing = [str(path) for path in normalized_bom_paths if not path.is_file()]
    if missing:
        raise FileNotFoundError("找不到 BOM 文件：\n" + "\n".join(missing))
    if not normalized_template_path.is_file():
        raise FileNotFoundError(f"找不到基础数据管制表模板：{normalized_template_path}")
    if normalized_output_path in [*normalized_bom_paths, normalized_template_path]:
        raise ValueError("输出文件不能覆盖 BOM 文件或模板文件。")

    _check_conversion_cancel(cancel_requested)
    records = read_and_merge_boms(
        normalized_bom_paths,
        progress_callback=progress_callback,
        cancel_requested=cancel_requested,
    )
    _check_conversion_cancel(cancel_requested)
    _emit_progress(progress_callback, 38, "正在读取基础数据管制表模板")
    template_book = openpyxl.load_workbook(normalized_template_path)
    template_sheet = template_book.active

    slot_starts, summary_start = ensure_record_capacity(template_sheet, len(records))
    clear_product_area(template_sheet, slot_starts)

    if len(slot_starts) < len(records):
        raise RuntimeError("模板记录槽位不足，无法写入全部 BOM 记录。")

    # 模板会预留空白记录槽位；只使用与 BOM 条目数相同的前 N 个槽位。
    target_slots = slot_starts[: len(records)]

    for index, (record, target_row) in enumerate(zip(records, target_slots), 1):
        _check_conversion_cancel(cancel_requested)
        write_record(
            template_sheet,
            target_row,
            record,
            record.get("_source_image"),
            geely_mode=geely_mode,
        )
        _emit_progress(
            progress_callback,
            40 + round(50 * index / len(records)),
            f"正在写入产品记录：{index}/{len(records)}",
        )

    _check_conversion_cancel(cancel_requested)
    _emit_progress(progress_callback, 92, "正在生成总成重量与汇总公式")
    write_assembly_mass_formulas(template_sheet, records, target_slots)
    update_summary_formula_ranges(template_sheet, summary_start)
    write_summary_formulas(template_sheet, summary_start, records, target_slots)
    template_book.calculation.fullCalcOnLoad = True
    template_book.calculation.forceFullCalc = True

    _check_conversion_cancel(cancel_requested)
    _emit_progress(progress_callback, 97, "正在保存合并后的基础数据管制表")
    normalized_output_path.parent.mkdir(parents=True, exist_ok=True)
    template_book.save(normalized_output_path)
    _emit_progress(progress_callback, 100, "模板合并转换完成")


def convert(bom_path: Path, template_path: Path, output_path: Path) -> None:
    """兼容旧调用：将一份 BOM 写入模板。"""
    convert_many([bom_path], template_path, output_path)


def default_output_path(bom_path: Path) -> Path:
    stem = bom_path.stem.replace("CATIA_图片清单_", "")
    return bom_path.parent / f"{stem}_基础数据管制表.xlsx"


def choose_paths_in_dialog(start_directory: Path) -> tuple[Path, Path, Path]:
    """显示文件选择窗口，依次获取 BOM、模板和输出位置。"""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ModuleNotFoundError as error:
        raise RuntimeError("当前 Python 未安装 tkinter，无法显示文件选择窗口。") from error

    root = None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        excel_files = [("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")]

        bom_name = filedialog.askopenfilename(
            title="第 1 步：选择 CATIA BOM 图片清单",
            initialdir=str(start_directory),
            filetypes=excel_files,
        )
        if not bom_name:
            raise RuntimeError("已取消选择 BOM 文件。")
        bom_path = Path(bom_name)

        template_name = filedialog.askopenfilename(
            title="第 2 步：选择基础数据管制表模板",
            initialdir=str(bom_path.parent),
            filetypes=excel_files,
        )
        if not template_name:
            raise RuntimeError("已取消选择模板文件。")
        template_path = Path(template_name)

        output_name = filedialog.asksaveasfilename(
            title="第 3 步：选择转换结果的保存位置",
            initialdir=str(bom_path.parent),
            initialfile=default_output_path(bom_path).name,
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not output_name:
            raise RuntimeError("已取消选择输出位置。")
        return bom_path, template_path, Path(output_name)
    except tk.TclError as error:
        raise RuntimeError(
            "无法显示文件选择窗口。请在 Windows 桌面会话中运行，"
            "或使用 --bom、--template、--output 参数指定文件。"
        ) from error
    finally:
        if root is not None:
            root.destroy()


def parse_args() -> argparse.Namespace:
    here = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description="将 CATIA 图片清单写入基础数据管制表产品区 A:AN")
    parser.add_argument(
        "--bom",
        type=Path,
        default=None,
        help="CATIA 图片清单 xlsx 文件",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=None,
        help="基础数据管制表模板 xlsx 文件",
    )
    parser.add_argument("--output", type=Path, default=None, help="输出 xlsx 文件")
    args = parser.parse_args()
    args.default_directory = here
    return args


def main() -> int:
    args = parse_args()
    if args.bom is None and args.template is None and args.output is None:
        bom_path, template_path, output_path = choose_paths_in_dialog(args.default_directory)
    else:
        if args.bom is None or args.template is None:
            raise ValueError("使用命令行模式时，必须同时提供 --bom 和 --template。")
        bom_path = args.bom
        template_path = args.template
        output_path = args.output or default_output_path(bom_path)

    bom_path = bom_path.resolve()
    template_path = template_path.resolve()
    output_path = output_path.resolve()

    if not bom_path.is_file():
        raise FileNotFoundError(f"未找到 BOM 文件：{bom_path}")
    if not template_path.is_file():
        raise FileNotFoundError(f"未找到模板文件：{template_path}")
    if output_path in (bom_path, template_path):
        raise ValueError("输出文件不能覆盖 BOM 文件或模板文件。")

    convert(bom_path, template_path, output_path)
    print(f"转换完成：{output_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"转换失败：{exc}", file=sys.stderr)
        raise SystemExit(1)
