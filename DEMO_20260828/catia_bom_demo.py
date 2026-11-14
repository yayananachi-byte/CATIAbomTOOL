# -*- coding: utf-8 -*-
"""CATIA BOM 导出 DEMO：独立输出，含质量、材质、厚度、包络尺寸和截图。

使用前请在 CATIA V5 中打开并激活一个 .CATProduct，然后运行本文件。
所有本次测试产生的 Excel、图片和日志均保存在本 DEMO 目录的“输出”文件夹。
"""

from __future__ import annotations

import sys
import time
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Optional, Tuple, Union

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


DEMO_ROOT = Path(__file__).resolve().parent
PROJECT_ROOT = DEMO_ROOT.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import catia_bom_export_v2 as bom


# 让底层截图模块也使用 DEMO 自己的输出目录，不覆盖主工程的历史结果。
OUTPUT_ROOT = DEMO_ROOT / "输出"
BOM_DIR = OUTPUT_ROOT / "BOM表"
IMAGE_DIR = OUTPUT_ROOT / "零件图片"
LOG_DIR = OUTPUT_ROOT / "日志"
bom.EXPORT_ROOT = OUTPUT_ROOT
bom.BOM_OUTPUT_DIR = BOM_DIR
bom.OUTPUT_DIR = BOM_DIR
bom.IMAGE_DIR = IMAGE_DIR
bom.LOG_DIR = LOG_DIR


@dataclass
class MassResult:
    value_kg: Optional[float] = None
    source: str = "未测量"
    status: str = "未测量"


def _read_mass_from_product(product: Any) -> Optional[float]:
    """优先采用 CATIA 产品惯性对象；CATIA Automation 的单位为 kg。"""
    if product is None:
        return None
    try:
        inertia = product.GetTechnologicalObject("Inertia")
        value = float(inertia.Mass)
        return round(value, 6) if value >= 0 else None
    except Exception:
        return None


def measure_part_mass(item: bom.BomItem) -> MassResult:
    """读取引用件质量，产品对象不可用时回退到 PartBody Inertia。"""
    representative = item.representative
    product_candidates = [representative.product, bom._reference_product(representative.product)]
    if representative.part_doc is not None:
        try:
            product_candidates.append(representative.part_doc.Product)
        except Exception:
            pass
    for product in product_candidates:
        value = _read_mass_from_product(product)
        if value is not None:
            return MassResult(value, "CATIA 产品 Inertia", "成功")

    part_doc = representative.part_doc
    inertia = None
    inertias = None
    try:
        if part_doc is None:
            raise bom.BomExportError("无法取得 CATPart 文档")
        part_doc.Activate()
        spa = part_doc.GetWorkbench("SPAWorkbench")
        inertias = spa.Inertias
        inertia = inertias.Add(bom._find_measurement_body(part_doc.Part))
        value = float(inertia.Mass)
        if value < 0:
            raise bom.BomExportError(f"CATIA 返回了无效质量：{value}")
        return MassResult(round(value, 6), "PartBody Inertia（后备）", "成功")
    except Exception as exc:
        return MassResult(None, "CATIA Inertia", f"失败：{exc}")
    finally:
        if inertia is not None and inertias is not None:
            try:
                name = bom.com_text(inertia.Name)
                if name:
                    inertias.Remove(name)
            except Exception:
                pass


def populate_masses(items, product_doc, log) -> dict[str, MassResult]:
    results: dict[str, MassResult] = {}
    for item in items.values():
        log.info("读取质量", item.part_number)
        try:
            result = measure_part_mass(item)
            results[item.key] = result
            if result.status == "成功":
                log.info(f"质量：{result.value_kg:.6f} kg（{result.source}）", item.part_number)
            else:
                log.warning(result.status, item.part_number)
        finally:
            if product_doc is not None:
                try:
                    product_doc.Activate()
                except Exception as exc:
                    log.warning(f"未能恢复总成文档：{exc}", item.part_number)
    return results


def _style(cell, font, fill, alignment, border) -> None:
    cell.font = font
    cell.fill = fill
    cell.alignment = alignment
    cell.border = border


def _setup_sheet(sheet, headers: list[str], title: str, widths: list[float]):
    thin = Side(style="thin", color="B7C9D6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    title_font = Font(name="微软雅黑", bold=True, size=13, color="FFFFFF")
    header_font = Font(name="微软雅黑", bold=True, size=10, color="FFFFFF")
    title_fill = PatternFill("solid", fgColor="1F4E78")
    header_fill = PatternFill("solid", fgColor="4472C4")
    last_column = get_column_letter(len(headers))
    sheet.merge_cells(f"A1:{last_column}1")
    for column in range(1, len(headers) + 1):
        _style(sheet.cell(1, column), title_font, title_fill, center, border)
    sheet["A1"] = title
    sheet.row_dimensions[1].height = 28
    for column, header in enumerate(headers, 1):
        _style(sheet.cell(2, column, header), header_font, header_fill, center, border)
        sheet.column_dimensions[get_column_letter(column)].width = widths[column - 1]
    sheet.row_dimensions[2].height = 32
    sheet.freeze_panes = "A3"
    return border, Font(name="微软雅黑", size=10), PatternFill("solid", fgColor="FFFFFF")


def _mass_values(mass: Optional[MassResult]) -> Tuple[Union[float, str], str, str]:
    if mass is None:
        return "", "", "未测量"
    return (mass.value_kg if mass.status == "成功" else "", mass.source, mass.status)


def write_demo_workbook(nodes, items, masses: dict[str, MassResult], output: Path, root_name: str, log) -> None:
    workbook = Workbook()
    tree_sheet = workbook.active
    tree_sheet.title = "产品结构"
    bom_sheet = workbook.create_sheet("BOM汇总")
    log_sheet = workbook.create_sheet("运行日志")
    border, data_font, white_fill = _setup_sheet(
        tree_sheet,
        ["树编号", "层级", "类型", "零件号", "版本", "零件名称", "实例名称", "父级路径",
         "属性材质", "PartBody材质", "材质比对", "质量(kg)", "质量来源", "质量状态",
         "包络长(mm)", "包络宽(mm)", "包络高(mm)", "包络来源", "包络状态",
         "名义厚度", "厚度来源", "本级数量", "缩略图", "截图状态"],
        f"CATIA BOM DEMO 产品结构：{root_name}",
        [14, 8, 12, 34, 10, 30, 32, 50, 28, 28, 14, 14, 24, 34, 16, 16, 16, 36, 42, 14, 22, 12, 24, 30],
    )
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    assembly_fill = PatternFill("solid", fgColor="D9EAF7")
    root_fill = PatternFill("solid", fgColor="9DC3E6")
    item_by_key = dict(items)
    for row, node in enumerate(nodes, 3):
        item = item_by_key.get(node.unique_key) if node.is_leaf else None
        material = item.material if item else None
        dimensions = item.dimensions if item else None
        thickness = item.nominal_thickness if item else None
        mass_value, mass_source, mass_status = _mass_values(masses.get(node.unique_key) if item else None)
        values = [
            node.tree_number, node.level, "根产品" if node.level == 0 else ("零件" if node.is_leaf else "子总成"),
            node.part_number, node.revision, ("　" * node.level) + node.nomenclature, node.instance_name, node.parent_path,
            material.property_material if material else "", material.body_material if material else "", material.comparison if material else "",
            mass_value, mass_source, mass_status,
            dimensions.length if dimensions and dimensions.status == "成功" else "",
            dimensions.width if dimensions and dimensions.status == "成功" else "",
            dimensions.height if dimensions and dimensions.status == "成功" else "",
            dimensions.source if dimensions else "", dimensions.status if dimensions else "",
            thickness.value if thickness else "", thickness.source if thickness else "", node.tree_quantity if item else "", "",
            item.image_status if item else node.image_status,
        ]
        fill = root_fill if node.level == 0 else (white_fill if node.is_leaf else assembly_fill)
        for column, value in enumerate(values, 1):
            _style(tree_sheet.cell(row, column, value), data_font, fill, left if column in (4, 6, 7, 8, 9, 10, 13, 14, 18, 19, 21, 24) else center, border)
        if item and item.image_path:
            size = bom.add_image(tree_sheet, f"W{row}", item.image_path)
            tree_sheet.row_dimensions[row].height = size[1] * 0.75 + 6 if size else 24
        else:
            tree_sheet.row_dimensions[row].height = 24
    if nodes:
        tree_sheet.auto_filter.ref = f"A2:X{len(nodes) + 2}"

    border, data_font, white_fill = _setup_sheet(
        bom_sheet,
        ["序号", "零件号", "版本", "零件名称", "数量", "来源文件", "属性材质", "PartBody材质",
         "质量(kg)", "质量来源", "质量状态", "包络长(mm)", "包络宽(mm)", "包络高(mm)",
         "包络来源", "包络状态", "名义厚度", "厚度来源", "缩略图", "截图状态"],
        f"CATIA BOM DEMO 汇总（唯一件 {len(items)} 种）",
        [8, 34, 10, 30, 8, 46, 28, 28, 14, 24, 34, 16, 16, 16, 36, 42, 14, 22, 24, 30],
    )
    alternate_fill = PatternFill("solid", fgColor="EAF2F8")
    for sequence, item in enumerate(items.values(), 1):
        row = sequence + 2
        mass_value, mass_source, mass_status = _mass_values(masses.get(item.key))
        values = [
            sequence, item.part_number, item.revision, item.nomenclature, item.quantity, item.source_file,
            item.material.property_material, item.material.body_material, mass_value, mass_source, mass_status,
            item.dimensions.length if item.dimensions.status == "成功" else "",
            item.dimensions.width if item.dimensions.status == "成功" else "",
            item.dimensions.height if item.dimensions.status == "成功" else "",
            item.dimensions.source, item.dimensions.status, item.nominal_thickness.value, item.nominal_thickness.source,
            "", item.image_status,
        ]
        fill = white_fill if sequence % 2 else alternate_fill
        for column, value in enumerate(values, 1):
            _style(bom_sheet.cell(row, column, value), data_font, fill, left if column in (2, 4, 6, 7, 8, 10, 11, 15, 16, 18, 20) else center, border)
        if item.image_path:
            size = bom.add_image(bom_sheet, f"S{row}", item.image_path)
            bom_sheet.row_dimensions[row].height = size[1] * 0.75 + 6 if size else 24
        else:
            bom_sheet.row_dimensions[row].height = 24
    if items:
        bom_sheet.auto_filter.ref = f"A2:T{len(items) + 2}"

    border, data_font, white_fill = _setup_sheet(log_sheet, ["级别", "对象", "信息"], "运行日志", [12, 40, 100])
    for row, (level, item_name, message) in enumerate(log.entries, 3):
        fill = PatternFill("solid", fgColor="FCE4D6") if level == "ERROR" else white_fill
        for column, value in enumerate((level, item_name, message), 1):
            _style(log_sheet.cell(row, column, value), data_font, fill, left, border)
        log_sheet.row_dimensions[row].height = 22
    if log.entries:
        log_sheet.auto_filter.ref = f"A2:C{len(log.entries) + 2}"
    workbook.save(output)


def main() -> int:
    bom.pythoncom.CoInitialize()
    log = bom.RunLog()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    started = time.time()
    try:
        for directory in (BOM_DIR, IMAGE_DIR, LOG_DIR):
            directory.mkdir(parents=True, exist_ok=True)
        catia = bom.connect_catia(log)
        product_doc, root_product = bom.get_active_product_document(catia)
        root_name = bom.com_text(root_product.PartNumber) or bom.com_text(product_doc.Name)
        log.info("遍历产品树", root_name)
        nodes = bom.traverse_product_tree(catia, root_product, log)
        items = bom.create_bom_items(nodes)
        log.info(f"遍历完成：{len(nodes)} 个树节点，{len(items)} 种唯一零件")
        if not items:
            raise bom.BomExportError("产品树没有可导出的叶子零件。")

        log.info("读取属性材质与 PartBody 材质")
        bom.populate_materials(items, log)
        log.info("读取名义厚度")
        bom.populate_nominal_thicknesses(items, log)
        masses = populate_masses(items, product_doc, log)
        log.info("测量包装包络（惯性主轴 Rough Stock；失败回退零件 XYZ）")
        bom.populate_rough_stock_dimensions(items, product_doc, log)

        for item in items.values():
            if item.representative.part_doc is not None:
                bom.hide_stale_measurement_geometry(item.representative.part_doc, log, item.part_number)
        assembly_nodes = [node for node in nodes if not node.is_leaf]
        for sequence, node in enumerate(assembly_nodes, 1):
            log.info("生成总成截图", f"{sequence}/{len(assembly_nodes)} {node.part_number}")
            bom.capture_assembly_image(catia, node, product_doc, log)
        for sequence, item in enumerate(items.values(), 1):
            log.info("生成唯一件截图", f"{sequence}/{len(items)} {item.part_number}")
            bom.capture_part_image(catia, item, product_doc, log)

        output = BOM_DIR / f"CATIA_BOM_DEMO_{bom.clean_filename(root_name)}_{timestamp}.xlsx"
        write_demo_workbook(nodes, items, masses, output, root_name, log)
        log.info(f"Excel 已生成：{output}")
        print(f"\nDEMO 完成，用时 {time.time() - started:.1f} 秒。")
        print(f"Excel：{output}")
        print(f"图片：{IMAGE_DIR}")
        return 0
    except bom.BomExportError as exc:
        log.error(str(exc))
        print(f"\nDEMO 失败：{exc}")
        return 1
    except Exception as exc:
        log.error(f"未处理异常：{exc}")
        traceback.print_exc()
        print(f"\nDEMO 失败：{exc}")
        return 1
    finally:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        log_path = LOG_DIR / f"CATIA_BOM_DEMO_日志_{timestamp}.txt"
        log.save(log_path)
        print(f"日志：{log_path}")
        bom.pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
