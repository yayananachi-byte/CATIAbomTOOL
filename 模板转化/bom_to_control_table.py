"""将 CATIA 图片清单 BOM 写入“基础数据管制表”的产品区（A:AN）。

依赖：openpyxl、Pillow（均随 Codex 运行环境提供）。

示例：
    python bom_to_control_table.py \
      --bom CATIA_图片清单_STU-0000509373_03_20260901_135340.xlsx \
      --template 客户XXX-项目名XXX-基础数据管制表-日期.xlsx

脚本只会写入产品区 A:AN；AO:DH 的模具、检具、夹具、成本等内容保持模板原样。
当 BOM 记录数超过模板现有的 16 个记录槽位时，脚本会在汇总区前按每项 6 行扩展记录区。
"""

from __future__ import annotations

import argparse
import re
import sys
from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange


DATA_START_ROW = 10
ROWS_PER_RECORD = 6
PRODUCT_LAST_COL = 40  # AN


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
    text = clean_text(value)
    if not text:
        return None
    match = re.search(r"(?<![A-Za-z])(-?\d+(?:\.\d+)?)\s*mm\b", text, flags=re.I)
    return float(match.group(1)) if match else None


def get_value(record: dict[str, Any], header: str) -> Any:
    value = record.get(header)
    if isinstance(value, str):
        return clean_text(value)
    return value


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


def root_and_parent_maps(
    records: list[dict[str, Any]]
) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
    by_tree = {record["_tree_no"]: record for record in records}
    roots: dict[str, dict[str, Any]] = {}
    for record in records:
        root_key = record["_path"][0]
        roots.setdefault(root_key, record)
    return by_tree, roots


def nearest_subassembly_number(record: dict[str, Any], by_tree: dict[str, dict[str, Any]]) -> str | None:
    """返回当前记录所属的最近分总成号；直属于根总成的零件返回空。"""
    path = record["_path"]
    for length in range(len(path), 1, -1):
        candidate = by_tree.get(".".join(path[:length]))
        if candidate and clean_text(candidate.get("类型")) == "子总成":
            return clean_text(candidate.get("零件号"))
    return None


def as_number(value: Any) -> float | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def aggregate_masses(records: list[dict[str, Any]]) -> dict[str, float | None]:
    """获取每个节点的本级总质量；总成/分总成无质量时递归汇总直属下级。"""
    children: dict[str, list[dict[str, Any]]] = {}
    for record in records:
        parent = ".".join(record["_path"][:-1])
        if parent:
            children.setdefault(parent, []).append(record)

    masses: dict[str, float | None] = {}
    for record in sorted(records, key=lambda item: len(item["_path"]), reverse=True):
        tree_no = record["_tree_no"]
        reported = as_number(get_value(record, "本级总质量(kg)"))
        if reported is not None:
            masses[tree_no] = reported
            continue
        child_masses = [
            masses.get(child["_tree_no"])
            for child in children.get(tree_no, [])
        ]
        numeric_child_masses = [mass for mass in child_masses if mass is not None]
        masses[tree_no] = sum(numeric_child_masses) if numeric_child_masses else None
    return masses


def is_standard_part(record: dict[str, Any]) -> bool:
    """CATIA BOM 中 Q/X 开头的零件号通常为紧固件等标准件。"""
    part_no = clean_text(record.get("零件号")) or ""
    return bool(re.match(r"^[QX]\d", part_no, flags=re.I))


def set_value(sheet, row: int, column: str, value: Any) -> None:
    if value is not None and value != "":
        sheet[f"{column}{row}"].value = value


def copy_thumbnail(sheet, source_image: ExcelImage | None, target_cell: str) -> None:
    if source_image is None:
        return
    image_bytes = source_image._data()
    target_image = ExcelImage(BytesIO(image_bytes))
    target_image.width = 160
    target_image.height = 77
    sheet.add_image(target_image, target_cell)


def write_record(
    sheet,
    target_row: int,
    serial_number: int,
    record: dict[str, Any],
    by_tree: dict[str, dict[str, Any]],
    roots: dict[str, dict[str, Any]],
    masses: dict[str, float | None],
    source_image: ExcelImage | None,
) -> None:
    path = record["_path"]
    depth = len(path) - 1
    if depth > 6:
        raise ValueError(
            f"树编号 {record['_tree_no']} 的层级为 {depth}，超过模板支持的 6 层（E:J）。"
        )

    part_type = clean_text(record.get("类型"))
    part_no = clean_text(record.get("零件号"))
    material = clean_text(record.get("PartBody Material")) or clean_text(record.get("属性 Material"))
    thickness = extract_thickness(record.get("名义厚度"))
    piece_mass = get_value(record, "单件质量(kg)")
    current_total_mass = get_value(record, "本级总质量(kg)")
    root_record = roots[path[0]]
    root_mass = masses.get(root_record["_tree_no"])

    # D 是序号；E:J 分别对应 BOM 的第 1 至第 6 层。根产品由其序号表示。
    set_value(sheet, target_row, "D", serial_number)
    if depth > 0:
        hierarchy_column = get_column_letter(4 + depth)
        set_value(sheet, target_row, hierarchy_column, record["_tree_no"])

    # 仅写入用户指定的产品字段。
    set_value(sheet, target_row, "M", clean_text(root_record.get("零件号")))
    set_value(sheet, target_row, "N", nearest_subassembly_number(record, by_tree))
    set_value(sheet, target_row, "O", part_no)
    set_value(sheet, target_row, "Q", trim_part_name(record.get("零件名称")))
    set_value(sheet, target_row, "R", get_value(record, "本级数量"))

    # 质量字段遵循模板语义：总成仅写 U，子总成写 V，零件写 X。
    if depth == 0:
        set_value(sheet, target_row, "U", root_mass)
    if part_type == "子总成":
        set_value(sheet, target_row, "V", masses.get(record["_tree_no"]))
    if part_type == "零件":
        if is_standard_part(record):
            set_value(sheet, target_row, "W", piece_mass)
        else:
            set_value(sheet, target_row, "X", piece_mass)

    set_value(sheet, target_row, "Y", material)
    set_value(sheet, target_row, "Z", thickness)
    set_value(sheet, target_row, "AB", get_value(record, "包装包络长(mm)"))
    set_value(sheet, target_row, "AC", get_value(record, "包装包络宽(mm)"))
    set_value(sheet, target_row, "AD", get_value(record, "包装包络高(mm)"))
    set_value(sheet, target_row, "AL", clean_text(record.get("版本")))
    copy_thumbnail(sheet, source_image, f"AN{target_row}")


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


def convert(bom_path: Path, template_path: Path, output_path: Path) -> None:
    records, images_by_row = read_bom(bom_path)
    template_book = openpyxl.load_workbook(template_path)
    template_sheet = template_book.active

    slot_starts, summary_start = ensure_record_capacity(template_sheet, len(records))
    clear_product_area(template_sheet, slot_starts)
    by_tree, roots = root_and_parent_maps(records)
    masses = aggregate_masses(records)

    for serial_number, (record, target_row) in enumerate(
        zip(records, slot_starts, strict=True), start=1
    ):
        write_record(
            template_sheet,
            target_row,
            serial_number,
            record,
            by_tree,
            roots,
            masses,
            images_by_row.get(record["_source_row"]),
        )

    update_summary_formula_ranges(template_sheet, summary_start)
    template_book.calculation.fullCalcOnLoad = True
    template_book.calculation.forceFullCalc = True

    output_path.parent.mkdir(parents=True, exist_ok=True)
    template_book.save(output_path)


def default_output_path(bom_path: Path) -> Path:
    stem = bom_path.stem.replace("CATIA_图片清单_", "")
    return bom_path.parent / f"{stem}_基础数据管制表.xlsx"


def parse_args() -> argparse.Namespace:
    here = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description="将 CATIA 图片清单写入基础数据管制表产品区 A:AN")
    parser.add_argument(
        "--bom",
        type=Path,
        default=here / "CATIA_图片清单_STU-0000509373_03_20260901_135340.xlsx",
        help="CATIA 图片清单 xlsx 文件",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=here / "客户XXX-项目名XXX-基础数据管制表-日期.xlsx",
        help="基础数据管制表模板 xlsx 文件",
    )
    parser.add_argument("--output", type=Path, default=None, help="输出 xlsx 文件")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    bom_path = args.bom.resolve()
    template_path = args.template.resolve()
    output_path = args.output.resolve() if args.output else default_output_path(bom_path)

    if not bom_path.is_file():
        raise FileNotFoundError(f"未找到 BOM 文件：{bom_path}")
    if not template_path.is_file():
        raise FileNotFoundError(f"未找到模板文件：{template_path}")

    convert(bom_path, template_path, output_path)
    print(f"转换完成：{output_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"转换失败：{exc}", file=sys.stderr)
        raise SystemExit(1)
