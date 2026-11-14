# -*- coding: utf-8 -*-
"""
CATIA 产品属性提取工具（汽车总成件版）
通过 CATIA V5 COM 接口，按产品树结构层级，提取产品中每个零件/子产品的：
  - 层级路径（按树结构缩进展示）
  - 零件名称 (Nomenclature)
  - 零件号 (Part Number)
  - 版本号 (Revision)
  - 材质 (Material)
  - 厚度 (Thickness)
  - 重量 (Mass)
并将结果写入 Excel 文件。

使用前请确保：
  1. CATIA V5 已安装并正在运行，且已打开一个 .CATProduct 文件
  2. 已安装 Python 依赖：pip install pywin32 openpyxl
"""

import os
import sys
import time

try:
    import win32com.client
except ImportError:
    print("错误：未安装 pywin32，请运行 pip install pywin32")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    print("错误：未安装 openpyxl，请运行 pip install openpyxl")
    sys.exit(1)


# ======================== 配置区 ========================
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "CATIA_零件属性表.xlsx")
# ========================================================


def connect_catia():
    """连接到正在运行的 CATIA 应用程序。"""
    try:
        catia = win32com.client.Dispatch("CATIA.Application")
        print("成功连接到 CATIA 应用程序。")
        return catia
    except Exception as e:
        print(f"错误：无法连接到 CATIA，请确认 CATIA 已启动。详情：{e}")
        sys.exit(1)


def get_active_product(catia):
    """获取当前活动文档的根产品。"""
    try:
        doc = catia.ActiveDocument
        if doc.Name.lower().endswith(".catproduct"):
            product = doc.Product
            print(f"当前活动文档：{doc.Name}")
            return product, doc
        else:
            print("错误：当前活动文档不是 .CATProduct，请打开一个产品文件。")
            sys.exit(1)
    except Exception as e:
        print(f"错误：无法获取活动文档。详情：{e}")
        sys.exit(1)


def get_part_mass(product):
    """
    通过 GetTechnologicalObject("Inertia") 获取质量（kg）。
    """
    try:
        inertia = product.GetTechnologicalObject("Inertia")
        mass = inertia.Mass
        return round(mass, 6)
    except Exception:
        return None


def get_material_on_product(product):
    """通过 MaterialManager 获取产品级别上应用的材质名称。"""
    try:
        manager = product.GetItem("CATMatManagerVBExt")
        material = win32com.client.VARIANT(
            win32com.client.pythoncom.VT_DISPATCH, None
        )
        manager.GetMaterialOnProduct(product, material)
        if material is not None and material.Name:
            return material.Name
    except Exception:
        pass
    return None


def get_material_on_part(part_doc):
    """通过 MaterialManager 获取零件级别上应用的材质名称。"""
    try:
        part = part_doc.Part
        manager = part.GetItem("CATMatManagerVBExt")
        material = win32com.client.VARIANT(
            win32com.client.pythoncom.VT_DISPATCH, None
        )
        manager.GetMaterialOnPart(part, material)
        if material is not None and material.Name:
            return material.Name
    except Exception:
        pass
    return None


def get_thickness_from_parameters(product, part_doc=None):
    """
    尝试从零件参数中获取厚度信息。
    常见厚度参数名：Thickness, 厚度, Epaisseur 等。
    同时也尝试从钣金特征中获取。
    """
    thickness = None
    thickness_keywords = [
        "thickness", "Thickness", "THICKNESS",
        "厚度", "厚み",
        "Epaisseur", "epaisseur",
        "Sheet\\Metal\\Thickness",
        "Wall\\Thickness",
        "壁厚",
    ]

    # 尝试从零件参数中获取
    if part_doc is not None:
        try:
            part = part_doc.Part
            params = part.Parameters
            for i in range(1, params.Count + 1):
                param = params.Item(i)
                param_name = param.Name
                for kw in thickness_keywords:
                    if kw in param_name:
                        try:
                            val = param.Value
                            if isinstance(val, (int, float)):
                                thickness = round(val, 4)
                                break
                        except Exception:
                            pass
                if thickness is not None:
                    break
        except Exception:
            pass

    # 尝试从钣金特征中获取厚度
    if thickness is None and part_doc is not None:
        try:
            part = part_doc.Part
            bodies = part.Bodies
            for i in range(1, bodies.Count + 1):
                body = bodies.Item(i)
                shapes = body.Shapes
                for j in range(1, shapes.Count + 1):
                    shape = shapes.Item(j)
                    try:
                        if hasattr(shape, "Thickness"):
                            thickness = round(shape.Thickness, 4)
                            break
                    except Exception:
                        pass
                if thickness is not None:
                    break
        except Exception:
            pass

    # 尝试从产品参数中获取
    if thickness is None:
        try:
            params = product.Parameters
            for i in range(1, params.Count + 1):
                param = params.Item(i)
                param_name = param.Name
                for kw in thickness_keywords:
                    if kw in param_name:
                        try:
                            val = param.Value
                            if isinstance(val, (int, float)):
                                thickness = round(val, 4)
                                break
                        except Exception:
                            pass
                if thickness is not None:
                    break
        except Exception:
            pass

    return thickness


def get_revision(product):
    """
    获取产品的版本号/修订号。
    优先从 Product.Revision 获取，若无则尝试从参数中查找。
    """
    revision = ""
    try:
        rev = product.Revision
        if rev and str(rev).strip():
            revision = str(rev).strip()
    except Exception:
        pass

    if not revision:
        try:
            params = product.Parameters
            rev_keywords = ["Revision", "revision", "版本", "Version", "版本号"]
            for i in range(1, params.Count + 1):
                param = params.Item(i)
                param_name = param.Name
                for kw in rev_keywords:
                    if kw in param_name:
                        try:
                            val = str(param.Value).strip()
                            if val:
                                revision = val
                                break
                        except Exception:
                            pass
                if revision:
                    break
        except Exception:
            pass

    return revision


def get_nomenclature(product):
    """
    获取零件名称（Nomenclature）。
    Nomenclature 是 CATIA 中用于描述零件功能名称的属性，
    例如 "支架"、"螺栓"、"外壳" 等。
    """
    nomenclature = ""
    try:
        nom = product.Nomenclature
        if nom and str(nom).strip():
            nomenclature = str(nom).strip()
    except Exception:
        pass

    # 如果 Nomenclature 为空，尝试从 DescriptionInst 获取
    if not nomenclature:
        try:
            desc = product.DescriptionInst
            if desc and str(desc).strip():
                nomenclature = str(desc).strip()
        except Exception:
            pass

    return nomenclature


def find_part_document(catia, part_number):
    """在 CATIA 已打开的文档中查找匹配零件号的 PartDocument。"""
    try:
        for d in range(1, catia.Documents.Count + 1):
            d_item = catia.Documents.Item(d)
            doc_name = d_item.Name
            # 匹配零件号（文档名通常为 "PartNumber.CATPart"）
            pn_lower = part_number.lower()
            dn_lower = doc_name.lower()
            if dn_lower.startswith(pn_lower) and dn_lower.endswith(".catpart"):
                if hasattr(d_item, "Part"):
                    return d_item
            # 也尝试直接匹配文档名
            if dn_lower == pn_lower + ".catpart":
                if hasattr(d_item, "Part"):
                    return d_item
    except Exception:
        pass
    return None


def is_leaf_part(product):
    """判断一个 Product 是否为叶子节点（最底层零件，无子产品）。"""
    try:
        return product.Products.Count == 0
    except Exception:
        return True


def traverse_and_collect(catia, root_product):
    """
    递归遍历产品树，按层级收集所有节点信息。
    返回列表，每个元素为一个字典，包含层级、路径、类型和各属性。
    """
    results = []

    # 先加入根产品节点
    root_info = _extract_node_info(catia, root_product, level=0,
                                   parent_path="", node_number="1",
                                   is_root=True)
    results.append(root_info)

    # 递归遍历子节点
    _traverse(catia, root_product, results, level=1,
              parent_path=root_product.PartNumber, parent_number="1")

    return results


def _traverse(catia, product, results, level, parent_path, parent_number):
    """递归遍历产品结构的子节点。"""
    products = product.Products
    if products.Count == 0:
        return

    for i in range(1, products.Count + 1):
        child = products.Item(i)
        node_number = f"{parent_number}.{i}"
        child_path = f"{parent_path} > {child.PartNumber}"

        # 提取节点信息
        node_info = _extract_node_info(catia, child, level=level,
                                       parent_path=parent_path,
                                       node_number=node_number,
                                       is_root=False)
        results.append(node_info)

        # 递归处理子产品
        if child.Products.Count > 0:
            _traverse(catia, child, results, level + 1,
                      parent_path=child_path,
                      parent_number=node_number)


def _extract_node_info(catia, product, level, parent_path, node_number, is_root):
    """提取单个节点的所有属性信息。"""
    instance_name = product.Name
    part_number = product.PartNumber
    nomenclature = get_nomenclature(product)
    revision = get_revision(product)
    leaf = is_leaf_part(product)

    # 材质：先尝试产品级
    material = get_material_on_product(product)

    # 厚度和零件级材质
    thickness = None
    part_doc = None

    if leaf:
        # 是叶子零件，尝试找 PartDocument
        part_doc = find_part_document(catia, part_number)

        if part_doc is not None:
            # 零件级材质
            if not material:
                mat_part = get_material_on_part(part_doc)
                if mat_part:
                    material = mat_part

            # 零件级厚度
            thickness = get_thickness_from_parameters(product, part_doc)

    # 如果还没有厚度，从产品参数获取
    if thickness is None:
        thickness = get_thickness_from_parameters(product, None)

    # 重量：只有叶子零件才获取重量
    mass = None
    if leaf:
        mass = get_part_mass(product)

    return {
        "level": level,
        "node_number": node_number,
        "instance_name": instance_name,
        "part_number": part_number,
        "nomenclature": nomenclature,
        "revision": revision,
        "material": material if material else "",
        "thickness": thickness if thickness is not None else "",
        "mass": mass if mass is not None else "",
        "is_leaf": leaf,
        "parent_path": parent_path,
        "is_root": is_root,
    }


def write_to_excel(results, output_path):
    """将结果写入 Excel 文件，按产品树结构展示。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "产品结构"

    # ---- 表头 ----
    headers = [
        "层级编号",        # 如 1, 1.1, 1.1.1
        "层级",           # 数字 0, 1, 2...
        "零件名称",       # Nomenclature
        "零件号",         # PartNumber
        "实例名称",       # InstanceName
        "版本号",         # Revision
        "材质",           # Material
        "厚度 (mm)",      # Thickness
        "重量 (kg)",      # Mass
        "所属总成",       # 父级路径
        "类型",           # 零件/子总成
    ]

    # 表头样式
    header_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # ---- 数据样式 ----
    # 总成行（子产品）样式 - 蓝色底
    assembly_font = Font(name="微软雅黑", bold=True, size=10, color="1F4E79")
    assembly_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")

    # 根产品样式 - 深蓝底
    root_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    root_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")

    # 叶子零件样式 - 白底
    part_font = Font(name="微软雅黑", size=10)
    part_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    # 缩进层级对应的浅色底色，增加视觉层次
    level_fills = [
        PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid"),  # 0 - 根
        PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"),  # 1 - 一级总成
        PatternFill(start_color="E9EFF7", end_color="E9EFF7", fill_type="solid"),  # 2 - 二级总成
        PatternFill(start_color="F2F7FC", end_color="F2F7FC", fill_type="solid"),  # 3 - 三级
        PatternFill(start_color="FAFCFE", end_color="FAFCFE", fill_type="solid"),  # 4+
    ]

    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center", indent=0)

    # ---- 填充数据 ----
    for row_idx, item in enumerate(results, 2):
        level = item["level"]
        is_leaf = item["is_leaf"]
        is_root = item["is_root"]

        # 根据层级生成缩进显示名
        indent = "  " * level
        display_part_number = f"{indent}{item['part_number']}"
        display_nomenclature = f"{indent}{item['nomenclature']}" if item['nomenclature'] else ""

        # 类型标识
        if is_root:
            type_label = "根产品"
        elif is_leaf:
            type_label = "零件"
        else:
            type_label = "子总成"

        row_data = [
            item["node_number"],
            level,
            display_nomenclature,
            display_part_number,
            item["instance_name"],
            item["revision"],
            item["material"],
            item["thickness"],
            item["mass"],
            item["parent_path"],
            type_label,
        ]

        # 选择行样式
        if is_root:
            row_font = root_font
            row_fill = level_fills[0]
        elif not is_leaf:
            # 子总成
            fill_idx = min(level, len(level_fills) - 1)
            row_font = assembly_font
            row_fill = level_fills[fill_idx]
        else:
            # 叶子零件
            fill_idx = min(level, len(level_fills) - 1)
            row_font = part_font
            row_fill = level_fills[fill_idx]

        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = row_font
            cell.fill = row_fill
            cell.border = thin_border
            # 零件名称和零件号列左对齐，其余居中
            if col_idx in (3, 4, 10):
                cell.alignment = left_align
            else:
                cell.alignment = center_align

    # ---- 列宽 ----
    col_widths = [12, 6, 25, 30, 30, 10, 18, 14, 14, 40, 10]
    for idx, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    # 冻结首行
    ws.freeze_panes = "A2"

    # 自动筛选
    if results:
        last_col = get_column_letter(len(headers))
        ws.auto_filter.ref = f"A1:{last_col}{len(results) + 1}"

    # ---- 添加汇总 Sheet ----
    _add_summary_sheet(wb, results)

    wb.save(output_path)
    print(f"结果已保存到：{output_path}")


def _add_summary_sheet(wb, results):
    """添加一个仅包含叶子零件的汇总表，方便筛选查看。"""
    ws = wb.create_sheet(title="零件汇总")

    headers = [
        "序号",
        "所属总成",
        "零件名称",
        "零件号",
        "版本号",
        "材质",
        "厚度 (mm)",
        "重量 (kg)",
    ]

    header_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    data_font = Font(name="微软雅黑", size=10)
    data_alignment = Alignment(horizontal="center", vertical="center")
    left_alignment = Alignment(horizontal="left", vertical="center")

    # 仅筛选叶子零件
    leaf_items = [item for item in results if item["is_leaf"]]
    seq = 0
    for item in leaf_items:
        seq += 1
        row_data = [
            seq,
            item["parent_path"],
            item["nomenclature"],
            item["part_number"],
            item["revision"],
            item["material"],
            item["thickness"],
            item["mass"],
        ]
        row_idx = seq + 1
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in (2, 3, 4):
                cell.alignment = left_alignment
            else:
                cell.alignment = data_alignment

    col_widths = [8, 40, 25, 30, 10, 18, 14, 14]
    for idx, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    ws.freeze_panes = "A2"
    if leaf_items:
        last_col = get_column_letter(len(headers))
        ws.auto_filter.ref = f"A1:{last_col}{len(leaf_items) + 1}"


def main():
    print("=" * 60)
    print("  CATIA 产品零件属性提取工具（汽车总成件版）")
    print("=" * 60)
    print()

    # 连接 CATIA
    catia = connect_catia()

    # 获取活动产品
    root_product, doc = get_active_product(catia)
    print(f"根产品：{root_product.PartNumber}")
    print(f"根产品名称：{get_nomenclature(root_product)}")
    print()

    # 收集零件信息
    print("正在遍历产品结构，提取零件属性...")
    start_time = time.time()

    results = traverse_and_collect(catia, root_product)

    elapsed = time.time() - start_time

    # 统计
    total_nodes = len(results)
    leaf_count = sum(1 for r in results if r["is_leaf"])
    assembly_count = sum(1 for r in results if not r["is_leaf"] and not r["is_root"])

    print(f"遍历完成，耗时 {elapsed:.2f} 秒。")
    print(f"  总节点数：{total_nodes}")
    print(f"  子总成数：{assembly_count}")
    print(f"  零件数量：{leaf_count}")

    if total_nodes <= 1:
        print("警告：产品结构中未找到任何子节点，请检查产品结构。")
        return

    # 写入 Excel
    write_to_excel(results, OUTPUT_FILE)
    print()
    print("提取完成！")
    print(f"  - 产品结构表：包含完整树结构层级")
    print(f"  - 零件汇总表：仅包含叶子零件，便于筛选")


if __name__ == "__main__":
    main()
