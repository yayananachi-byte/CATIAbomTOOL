# -*- coding: utf-8 -*-
"""CATIA V5：仅导出唯一零件截图与数量（不测量尺寸、不修改零件）。

运行前：在 CATIA 中打开并激活 .CATProduct。
运行：py .\catia_bom_screenshot.py
"""

from __future__ import annotations

import time
import traceback
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 复用已验证的 CATIA 连接、产品树去重、CaptureToFile 截图逻辑；不会调用尺寸模块。
from catia_bom_export_v2 import (
    BOM_OUTPUT_DIR,
    IMAGE_DIR,
    IMAGE_SIZE,
    LOG_DIR,
    BomExportError,
    RunLog,
    add_image,
    capture_assembly_image,
    capture_part_image,
    clean_filename,
    com_text,
    connect_catia,
    create_bom_items,
    get_active_product_document,
    hide_stale_measurement_geometry,
    populate_materials,
    populate_masses,
    populate_nominal_thicknesses,
    populate_rough_stock_dimensions,
    traverse_product_tree,
)


def _style(cell, font, fill, alignment, border) -> None:
    cell.font = font
    cell.fill = fill
    cell.alignment = alignment
    cell.border = border


def write_image_index(nodes, items, output: Path, root_name: str, log: RunLog) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "产品结构"

    thin = Side(style="thin", color="B7C9D6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    title_font = Font(name="微软雅黑", bold=True, size=13, color="FFFFFF")
    header_font = Font(name="微软雅黑", bold=True, size=10, color="FFFFFF")
    data_font = Font(name="微软雅黑", size=10)
    title_fill = PatternFill("solid", fgColor="1F4E78")
    header_fill = PatternFill("solid", fgColor="4472C4")
    white = PatternFill("solid", fgColor="FFFFFF")
    # 一个表中同时保存完整树结构、同件数量与缩略图。
    headers = [
        "树编号", "层级", "类型", "零件号", "版本", "零件名称", "实例名称", "父级路径",
        "属性 Material", "PartBody Material", "材质比对",
        "单件质量(kg)", "本级总质量(kg)", "测量密度(kg/m³)", "密度状态", "质量来源", "质量状态",
        "包装包络长(mm)", "包装包络宽(mm)", "包装包络高(mm)",
        "包络来源", "包络状态",
        "名义厚度", "厚度来源", "本级数量", "缩略图", "截图状态",
    ]
    widths = [14, 8, 12, 36, 10, 30, 34, 56, 34, 38, 16, 15, 17, 20, 34, 24, 34, 16, 16, 16, 38, 52, 16, 24, 14, 25, 34]
    last_column = get_column_letter(len(headers))
    sheet.merge_cells(f"A1:{last_column}1")
    _style(sheet["A1"], title_font, title_fill, center, border)
    sheet["A1"] = f"产品结构与零件图片：{root_name}"
    for column in range(2, len(headers) + 1):
        _style(sheet.cell(1, column), title_font, title_fill, center, border)
    sheet.row_dimensions[1].height = 28
    for column, header in enumerate(headers, 1):
        _style(sheet.cell(2, column, header), header_font, header_fill, center, border)
    sheet.row_dimensions[2].height = 30
    for column, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(column)].width = width

    item_by_key = dict(items)
    assembly_fill = PatternFill("solid", fgColor="D9EAF7")
    root_fill = PatternFill("solid", fgColor="9DC3E6")
    for row, node in enumerate(nodes, 3):
        item = item_by_key.get(node.unique_key) if node.is_leaf else None
        image_path = item.image_path if item else node.image_path
        image_status = item.image_status if item else node.image_status
        material = item.material if item else None
        thickness = item.nominal_thickness if item else None
        dimensions = item.dimensions if item else None
        mass = item.mass if item else None
        unit_mass = mass.value_kg if mass and mass.status == "成功" else ""
        level_mass = round(unit_mass * node.tree_quantity, 6) if unit_mass != "" else ""
        density = mass.density_kg_m3 if mass and mass.status == "成功" and mass.density_kg_m3 is not None else ""
        type_label = "根产品" if node.level == 0 else ("零件" if node.is_leaf else "子总成")
        values = [
            node.tree_number,
            node.level,
            type_label,
            node.part_number,
            node.revision,
            ("　" * node.level) + node.nomenclature,
            node.instance_name,
            node.parent_path,
            material.property_material if material else "",
            material.body_material if material else "",
            material.comparison if material else "",
            unit_mass,
            level_mass,
            density,
            mass.density_status if mass else "",
            mass.source if mass else "",
            mass.status if mass else "",
            dimensions.length if dimensions and dimensions.status == "成功" else "",
            dimensions.width if dimensions and dimensions.status == "成功" else "",
            dimensions.height if dimensions and dimensions.status == "成功" else "",
            dimensions.source if dimensions else "",
            dimensions.status if dimensions else "",
            thickness.value if thickness else "",
            thickness.source if thickness else "",
            node.tree_quantity if item else "",
            "",
            image_status,
        ]
        fill = root_fill if node.level == 0 else (white if node.is_leaf else assembly_fill)
        for column, value in enumerate(values, 1):
            _style(sheet.cell(row, column, value), data_font, fill, left if column in (4, 6, 7, 8, 9, 10, 15, 16, 17, 21, 22, 24, 27) else center, border)
        for column in (12, 13):
            sheet.cell(row, column).number_format = "0.000000"
        sheet.cell(row, 14).number_format = "#,##0.000"
        if material and material.comparison == "不一致":
            mismatch_fill = PatternFill("solid", fgColor="FFC7CE")
            mismatch_font = Font(name="微软雅黑", size=10, bold=True, color="9C0006")
            for column in (9, 10, 11):
                _style(sheet.cell(row, column), mismatch_font, mismatch_fill, left if column in (9, 10) else center, border)
        if mass and mass.status == "成功" and mass.density_status != "已读取（均匀密度）":
            density_warning_fill = PatternFill("solid", fgColor="FFF2CC")
            density_warning_font = Font(name="微软雅黑", size=10, bold=True, color="7F6000")
            for column in range(12, 18):
                _style(
                    sheet.cell(row, column),
                    density_warning_font,
                    density_warning_fill,
                    left if column in (15, 16, 17) else center,
                    border,
                )
        display_size = add_image(sheet, f"Z{row}", image_path)
        if display_size:
            sheet.row_dimensions[row].height = display_size[1] * 0.75 + 6
        else:
            sheet.row_dimensions[row].height = 24

    if nodes:
        sheet.auto_filter.ref = f"A2:{last_column}{len(nodes) + 2}"
    sheet.freeze_panes = "A3"

    workbook.save(output)


def main() -> None:
    log = RunLog()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    started = time.time()
    try:
        BOM_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        catia = connect_catia(log)
        product_doc, root_product = get_active_product_document(catia)
        root_name = com_text(root_product.PartNumber) or com_text(product_doc.Name)
        log.info("遍历产品树", root_name)
        nodes = traverse_product_tree(catia, root_product, log)
        items = create_bom_items(nodes)
        if not items:
            raise BomExportError("产品树中没有可截图的叶子零件。")

        log.info("读取属性 Material 与 PartBody Material")
        populate_materials(items, log)
        log.info("读取名义厚度（仅加厚曲面特征；不切换零件工作台）")
        populate_nominal_thicknesses(items, log)
        log.info("测量单件质量（CATPart PartBody Inertia；产品 Inertia 仅后备）")
        populate_masses(items, product_doc, log)
        log.info("测量包装包络（PartBody 惯性主轴 Rough Stock，失败时三视图投影后备）")
        populate_rough_stock_dimensions(items, product_doc, log)
        IMAGE_DIR.mkdir(parents=True, exist_ok=True)
        # 清理上一轮失败包络测量遗留的辅助点/面，必须在总成截图之前完成。
        for item in items.values():
            if item.representative.part_doc is not None:
                hide_stale_measurement_geometry(item.representative.part_doc, log, item.part_number)
        assembly_nodes = [node for node in nodes if not node.is_leaf]
        log.info(f"共找到 {len(assembly_nodes)} 个总成节点、{len(items)} 种唯一零件；开始截图")
        for index, node in enumerate(assembly_nodes, 1):
            log.info("生成总成截图", f"{index}/{len(assembly_nodes)} {node.tree_number} {node.part_number}")
            capture_assembly_image(catia, node, product_doc, log)
        for index, item in enumerate(items.values(), 1):
            log.info("生成唯一件截图", f"{index}/{len(items)} {item.part_number}")
            capture_part_image(catia, item, product_doc, log)

        output = BOM_OUTPUT_DIR / f"CATIA_图片清单_{clean_filename(root_name)}_{timestamp}.xlsx"
        write_image_index(nodes, items, output, root_name, log)
        log.info(f"图片清单已生成：{output}")
        print(f"\n截图完成，用时 {time.time() - started:.1f} 秒。")
        print(f"图片目录：{IMAGE_DIR}")
        print(f"图片清单：{output}")
    except BomExportError as exc:
        log.error(str(exc))
        print(f"\n截图失败：{exc}")
    except Exception as exc:
        log.error(f"未处理异常：{exc}")
        traceback.print_exc()
    finally:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        log_path = LOG_DIR / f"CATIA_图片日志_{timestamp}.txt"
        log.save(log_path)
        print(f"日志：{log_path}")


if __name__ == "__main__":
    main()
