# -*- coding: utf-8 -*-
"""
CATIA 产品属性提取工具（汽车总成件版）
通过 CATIA V5 COM 接口，按产品树结构层级，提取产品中每个零件/子产品的：
  - 零件缩略图（自动截图嵌入 Excel）
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
  2. 已安装 Python 依赖：pip install pywin32 openpyxl Pillow
"""

import os
import sys
import time
import re

try:
    import win32com.client
    import win32gui
    import win32ui
    from ctypes import windll
except ImportError:
    print("错误：未安装 pywin32，请运行 pip install pywin32")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XlImage
except ImportError:
    print("错误：未安装 openpyxl，请运行 pip install openpyxl")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("错误：未安装 Pillow，请运行 pip install Pillow")
    sys.exit(1)


# ======================== 配置区 ========================
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "CATIA_零件属性表.xlsx")
IMAGE_DIR = os.path.join(OUTPUT_DIR, "零件图片")
IMAGE_WIDTH = 200       # Excel 中图片宽度（像素）
IMAGE_HEIGHT = 150      # Excel 中图片高度（像素）
THUMB_WIDTH = 2560       # 导出图片/缩略图原始宽度
THUMB_HEIGHT = 1600      # 导出图片/缩略图原始高度
# ========================================================


def sanitize_filename(name):
    """将文件名中的非法字符替换为下划线。"""
    return re.sub(r'[\\/:*?"<>|]', '_', name)


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


def _clean_image_cache(image_dir):
    """清空图片缓存目录中的所有图片文件。"""
    if not os.path.exists(image_dir):
        return
    for f in os.listdir(image_dir):
        if f.lower().endswith(('.bmp', '.png', '.jpg', '.jpeg')):
            try:
                os.remove(os.path.join(image_dir, f))
            except Exception:
                pass


def _screenshot_catia_window(catia, save_path):
    """
    使用 Pillow ImageGrab 截取 CATIA 3D 视图区域。
    通过 Viewers 集合获取 3D 视图在窗口中的位置，只截取视图区域。
    """
    from PIL import ImageGrab

    # 获取 CATIA 窗口在屏幕上的位置
    try:
        catia_window = catia.ActiveWindow
        win_left = catia_window.Left   # 窗口左上角屏幕 X
        win_top = catia_window.Top     # 窗口左上角屏幕 Y
        win_width = catia_window.Width
        win_height = catia_window.Height
    except Exception:
        # 获取不到窗口信息就截全屏
        img = ImageGrab.grab()
        if img.mode == 'RGBA':
            img = img.convert('RGB')
        img.save(save_path, "PNG")
        return True

    # 尝试获取 3D 视图区域在窗口内的位置
    viewer_bbox = None
    try:
        viewers = catia_window.Viewers
        if viewers.Count > 0:
            # 取第一个 viewer（通常是主 3D 视图）
            v = viewers.Item(1)
            # viewer 的 Left/Top/Width/Height 是相对于窗口的坐标
            v_left = v.Left
            v_top = v.Top
            v_width = v.Width
            v_height = v.Height
            # 转换为屏幕坐标
            scr_left = win_left + v_left
            scr_top = win_top + v_top
            scr_right = scr_left + v_width
            scr_bottom = scr_top + v_height
            # 确保坐标有效
            if scr_right > scr_left and scr_bottom > scr_top:
                viewer_bbox = (int(scr_left), int(scr_top),
                               int(scr_right), int(scr_bottom))
    except Exception:
        pass

    # 截屏：优先截 3D 视图区域，获取不到则截整个窗口
    if viewer_bbox:
        img = ImageGrab.grab(bbox=viewer_bbox)
    else:
        bbox = (win_left, win_top, win_left + win_width, win_top + win_height)
        img = ImageGrab.grab(bbox=bbox)

    if img.mode == 'RGBA':
        img = img.convert('RGB')
    img.save(save_path, "PNG")
    return True


def capture_part_image(catia, product_doc, target_product, image_dir):
    """
    在总成视图中：隐藏其他所有零件 → 仅显示当前零件 → 屏幕截图。
    使用 Pillow ImageGrab 直接截取屏幕，最简单可靠。
    流程：
      1. 隐藏所有叶子零件
      2. 仅显示目标零件
      3. 刷新视图 + 适应窗口
      4. 取消选中状态（避免红色选中边框）
      5. ImageGrab 截取 CATIA 3D视图区域
      6. 恢复所有零件显示
      7. 生成缩略图
    返回缩略图的完整路径，失败返回 None。
    """
    safe_name = sanitize_filename(target_product.PartNumber)
    png_path = os.path.join(image_dir, f"{safe_name}.png")
    thumb_path = os.path.join(image_dir, f"{safe_name}_thumb.png")

    selection = product_doc.Selection
    vis_props = selection.VisProperties

    # 常量定义
    catVisPropertyNoShowAttr = 1  # 隐藏
    catVisPropertyShowAttr = 0    # 显示

    try:
        # ---- Step 1: 隐藏所有叶子零件 ----
        selection.Clear()
        root_product = product_doc.Product
        _select_all_leaf_products(root_product, selection)
        if selection.Count > 0:
            vis_props.SetShow(catVisPropertyNoShowAttr)
        time.sleep(0.5)

        # ---- Step 2: 显示目标零件 ----
        selection.Clear()
        try:
            selection.Add(target_product)
        except Exception:
            try:
                selection.Search(f"Name={target_product.Name},all")
            except Exception:
                pass

        if selection.Count > 0:
            vis_props.SetShow(catVisPropertyShowAttr)
        time.sleep(0.5)

        # ---- Step 3: 刷新视图并适应窗口 ----
        try:
            viewer = catia.ActiveWindow.ActiveViewer
            # 强制刷新视图，确保隐藏/显示操作已生效
            viewer.Update()
            time.sleep(0.5)
            viewer.Update()
            time.sleep(0.3)

            # FitAllIn（多语言尝试）
            for cmd_name in ["全部适应", "适应全部", "Fit All In", "Zoom to Fit"]:
                try:
                    catia.StartCommand(cmd_name)
                    time.sleep(0.5)
                    break
                except Exception:
                    continue

            # 再次强制刷新，确保画面完全更新
            viewer.Update()
            time.sleep(0.8)
        except Exception:
            pass

        # ---- Step 4: 取消选中状态（避免红色选中边框） ----
        selection.Clear()
        time.sleep(0.3)

        # ---- Step 5: 屏幕截图 ----
        captured = False
        try:
            captured = _screenshot_catia_window(catia, png_path)
        except Exception as e:
            print(f"    屏幕截图失败：{e}")

        # ---- Step 6: 恢复所有零件显示 ----
        selection.Clear()
        _select_all_leaf_products(root_product, selection)
        if selection.Count > 0:
            vis_props.SetShow(catVisPropertyShowAttr)
        selection.Clear()
        time.sleep(0.3)

        # ---- Step 7: 生成缩略图（不裁切，直接使用完整截图） ----
        if captured and os.path.exists(png_path):
            try:
                img = Image.open(png_path)
                if img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')

                # 保存完整截图
                img.save(png_path, "PNG")

                # 生成缩略图
                img_thumb = img.copy()
                img_thumb.thumbnail((THUMB_WIDTH, THUMB_HEIGHT), Image.LANCZOS)
                img_thumb.save(thumb_path, "PNG")

                print(f"    截图成功：{os.path.getsize(thumb_path)} bytes")
                return thumb_path
            except Exception as e:
                print(f"    图片处理失败：{e}")
        else:
            print(f"    截图失败")

    except Exception as e:
        print(f"    截图过程出错：{e}")
        # 确保恢复显示
        try:
            selection.Clear()
            _select_all_leaf_products(product_doc.Product, selection)
            if selection.Count > 0:
                vis_props.SetShow(catVisPropertyShowAttr)
            selection.Clear()
        except Exception:
            pass

    return None


def _select_all_leaf_products(root_product, selection):
    """
    递归选中产品树下所有叶子零件（用于批量隐藏/显示）。
    """
    try:
        products = root_product.Products
        for i in range(1, products.Count + 1):
            child = products.Item(i)
            if child.Products.Count == 0:
                # 叶子零件，添加到选择集
                selection.Add(child)
            else:
                # 子总成，递归
                _select_all_leaf_products(child, selection)
    except Exception:
        pass


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

    # 构建完整层级路径和祖先列表
    if is_root:
        full_path = part_number
        ancestors = [part_number]
    elif parent_path:
        full_path = f"{parent_path} > {part_number}"
        ancestors = full_path.split(" > ")
    else:
        full_path = part_number
        ancestors = [part_number]

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
        "full_path": full_path,
        "ancestors": ancestors,
        "is_root": is_root,
    }


def capture_all_images(catia, results, image_dir, product_doc, root_product):
    """
    为所有叶子零件生成截图。
    在总成视图中逐个隐藏其他零件、显示当前零件、截图。
    每次运行会先清空旧图片缓存，确保截图是最新的。
    返回 { part_number: thumbnail_path } 的字典。
    """
    # 创建图片目录
    os.makedirs(image_dir, exist_ok=True)

    # 清空旧图片缓存（避免上次错误的截图被复用）
    _clean_image_cache(image_dir)

    image_map = {}
    leaf_items = [r for r in results if r["is_leaf"]]
    total = len(leaf_items)

    if total == 0:
        return image_map

    print(f"\n正在为 {total} 个零件生成截图（隐藏其他件 → 显示当前件 → 截图）...")

    # 先构建 part_number -> Product 对象的映射
    pn_to_product = {}
    _collect_leaf_products(root_product, pn_to_product)

    for idx, item in enumerate(leaf_items, 1):
        pn = item["part_number"]
        print(f"  [{idx}/{total}] 截图：{pn}")

        target_product = pn_to_product.get(pn)
        if target_product is None:
            print(f"    未找到零件对象，跳过")
            continue

        thumb_path = capture_part_image(catia, product_doc, target_product, image_dir)
        if thumb_path:
            image_map[pn] = thumb_path
        else:
            print(f"    截图失败，跳过")

    # 确保所有零件恢复显示
    try:
        selection = product_doc.Selection
        vis_props = selection.VisProperties
        selection.Clear()
        _select_all_leaf_products(root_product, selection)
        if selection.Count > 0:
            vis_props.SetShow(0)  # catVisPropertyShowAttr = 0
        selection.Clear()
        # 刷新视图
        catia.ActiveWindow.ActiveViewer.Update()
    except Exception:
        pass

    success = len(image_map)
    print(f"截图完成：成功 {success}/{total}")

    return image_map


def _collect_leaf_products(product, pn_map):
    """
    递归收集产品树中所有叶子零件的 { PartNumber: Product } 映射。
    """
    try:
        products = product.Products
        for i in range(1, products.Count + 1):
            child = products.Item(i)
            if child.Products.Count == 0:
                pn_map[child.PartNumber] = child
            else:
                _collect_leaf_products(child, pn_map)
    except Exception:
        pass


def write_to_excel(results, output_path, image_map=None, product_name=""):
    """将结果写入 Excel 文件，按产品树结构展示，并嵌入零件图片。
    Row 1: 完整上级总成链（合并单元格）
    Row 2: 列表头（第一列为零件号）
    Row 3+: 数据行
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "产品结构"

    has_images = image_map and len(image_map) > 0

    # ---- 列表头（第一列为零件号） ----
    headers = [
        "零件号",         # A: PartNumber
        "层级编号",        # B: 如 1, 1.1, 1.1.1
        "缩略图",         # C: 零件图片
        "零件名称",       # D: Nomenclature
        "实例名称",       # E: InstanceName
        "版本号",         # F: Revision
        "材质",           # G: Material
        "厚度 (mm)",      # H: Thickness
        "重量 (kg)",      # I: Mass
        "类型",           # J: 零件/子总成
    ]

    num_cols = len(headers)
    last_col_letter = get_column_letter(num_cols)

    # ---- 样式定义 ----
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # 标题行样式（第一行：上级总成链）
    title_font = Font(name="微软雅黑", bold=True, size=13, color="FFFFFF")
    title_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    title_alignment = Alignment(horizontal="center", vertical="center")

    # 表头样式（第二行）
    header_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ---- Row 1: 完整上级总成链（合并单元格） ----
    # 从结果中找到最深的叶子节点，取其 full_path 作为完整上级总成链
    hierarchy_chain = product_name
    for r in results:
        if r["is_leaf"] and r.get("full_path"):
            hierarchy_chain = r["full_path"]
            break
    # 如果没找到叶子，用根产品
    if hierarchy_chain == product_name:
        for r in results:
            if r["is_root"]:
                hierarchy_chain = r["full_path"] if r.get("full_path") else r["part_number"]
                break

    ws.merge_cells(f"A1:{last_col_letter}1")
    title_cell = ws.cell(row=1, column=1, value=hierarchy_chain)
    title_cell.font = title_font
    title_cell.fill = title_fill
    title_cell.alignment = title_alignment
    title_cell.border = thin_border
    for col_idx in range(2, num_cols + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = title_fill
        c.border = thin_border
    ws.row_dimensions[1].height = 36

    # ---- Row 2: 列表头 ----
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    ws.row_dimensions[2].height = 28

    # ---- 数据样式 ----
    assembly_font = Font(name="微软雅黑", bold=True, size=10, color="1F4E79")
    root_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    part_font = Font(name="微软雅黑", size=10)

    level_fills = [
        PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid"),  # 0 - 根
        PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"),  # 1
        PatternFill(start_color="E9EFF7", end_color="E9EFF7", fill_type="solid"),  # 2
        PatternFill(start_color="F2F7FC", end_color="F2F7FC", fill_type="solid"),  # 3
        PatternFill(start_color="FAFCFE", end_color="FAFCFE", fill_type="solid"),  # 4+
    ]

    # 斑马纹交替行颜色（用于叶子零件行）
    zebra_fill_a = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # 白色
    zebra_fill_b = PatternFill(start_color="EDF2F9", end_color="EDF2F9", fill_type="solid")  # 淡蓝灰

    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center", indent=0)

    # 行高：含图片行需要更高
    row_height_normal = 20
    row_height_image = IMAGE_HEIGHT + 10

    # 斑马纹计数器（仅对叶子零件行计数）
    leaf_row_counter = 0

    # ---- 填充数据（从 Row 3 开始） ----
    for row_idx, item in enumerate(results, 3):
        level = item["level"]
        is_leaf = item["is_leaf"]
        is_root = item["is_root"]

        # 类型标识
        if is_root:
            type_label = "根产品"
        elif is_leaf:
            type_label = "零件"
        else:
            type_label = "子总成"

        # 缩进显示名称
        indent = "  " * level
        display_nomenclature = f"{indent}{item['nomenclature']}" if item['nomenclature'] else ""

        row_data = [
            item["part_number"],        # A: 零件号
            item["node_number"],        # B: 层级编号
            "",                         # C: 缩略图列
            display_nomenclature,       # D: 零件名称
            item["instance_name"],      # E: 实例名称
            item["revision"],           # F: 版本号
            item["material"],           # G: 材质
            item["thickness"],          # H: 厚度
            item["mass"],               # I: 重量
            type_label,                 # J: 类型
        ]

        # 选择行样式
        if is_root:
            row_font = root_font
            row_fill = level_fills[0]
        elif not is_leaf:
            # 子总成行：按层级着色
            fill_idx = min(level, len(level_fills) - 1)
            row_font = assembly_font
            row_fill = level_fills[fill_idx]
        else:
            # 叶子零件行：斑马纹交替
            if leaf_row_counter % 2 == 0:
                row_fill = zebra_fill_a
            else:
                row_fill = zebra_fill_b
            leaf_row_counter += 1
            row_font = part_font

        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = row_font
            cell.fill = row_fill
            cell.border = thin_border
            # 零件号、零件名称、实例名称左对齐
            if col_idx in (1, 4, 5):
                cell.alignment = left_align
            else:
                cell.alignment = center_align

        # 设置行高
        if is_leaf and has_images and item["part_number"] in image_map:
            ws.row_dimensions[row_idx].height = row_height_image
        else:
            ws.row_dimensions[row_idx].height = row_height_normal

    # ---- 插入图片 ----
    if has_images:
        print("正在将图片嵌入 Excel...")
        for row_idx, item in enumerate(results, 3):
            if item["is_leaf"]:
                pn = item["part_number"]
                if pn in image_map:
                    thumb_path = image_map[pn]
                    try:
                        img = XlImage(thumb_path)
                        img.width = IMAGE_WIDTH
                        img.height = IMAGE_HEIGHT
                        cell_ref = f"C{row_idx}"
                        ws.add_image(img, cell_ref)
                    except Exception as e:
                        print(f"  嵌入图片失败 ({pn})：{e}")

    # ---- 列宽 ----
    col_widths = [30, 10, 18, 25, 30, 10, 18, 14, 14, 10]
    for idx, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    # 冻结前两行（标题行 + 表头行）
    ws.freeze_panes = "A3"

    # 自动筛选（从表头行开始）
    if results:
        ws.auto_filter.ref = f"A2:{last_col_letter}{len(results) + 2}"

    # ---- 添加汇总 Sheet ----
    _add_summary_sheet(wb, results, image_map)

    wb.save(output_path)
    print(f"结果已保存到：{output_path}")


def _add_summary_sheet(wb, results, image_map=None):
    """添加一个仅包含叶子零件的汇总表，方便筛选查看。"""
    ws = wb.create_sheet(title="零件汇总")

    has_images = image_map and len(image_map) > 0

    headers = [
        "零件号",
        "序号",
        "缩略图",
        "零件名称",
        "所属总成",
        "版本号",
        "材质",
        "厚度 (mm)",
        "重量 (kg)",
    ]

    num_cols = len(headers)
    last_col_letter = get_column_letter(num_cols)

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # 标题行样式
    title_font = Font(name="微软雅黑", bold=True, size=13, color="FFFFFF")
    title_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    title_alignment = Alignment(horizontal="center", vertical="center")

    header_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Row 1: 标题行
    hierarchy_chain = ""
    for r in results:
        if r["is_leaf"] and r.get("full_path"):
            hierarchy_chain = r["full_path"]
            break
    if not hierarchy_chain:
        for r in results:
            if r["is_root"]:
                hierarchy_chain = r.get("full_path", r["part_number"])
                break

    ws.merge_cells(f"A1:{last_col_letter}1")
    title_cell = ws.cell(row=1, column=1, value=hierarchy_chain)
    title_cell.font = title_font
    title_cell.fill = title_fill
    title_cell.alignment = title_alignment
    title_cell.border = thin_border
    for col_idx in range(2, num_cols + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = title_fill
        c.border = thin_border
    ws.row_dimensions[1].height = 36

    # Row 2: 表头
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    ws.row_dimensions[2].height = 28

    data_font = Font(name="微软雅黑", size=10)
    data_alignment = Alignment(horizontal="center", vertical="center")
    left_alignment = Alignment(horizontal="left", vertical="center")

    # 斑马纹
    zebra_fill_a = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    zebra_fill_b = PatternFill(start_color="EDF2F9", end_color="EDF2F9", fill_type="solid")

    # 仅筛选叶子零件
    leaf_items = [item for item in results if item["is_leaf"]]
    seq = 0
    for item in leaf_items:
        seq += 1
        row_data = [
            item["part_number"],
            seq,
            "",  # 缩略图列先留空
            item["nomenclature"],
            item["parent_path"],
            item["revision"],
            item["material"],
            item["thickness"],
            item["mass"],
        ]
        row_idx = seq + 2  # 从 Row 3 开始
        row_fill = zebra_fill_a if seq % 2 == 1 else zebra_fill_b
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.fill = row_fill
            cell.border = thin_border
            if col_idx in (1, 4, 5):  # 零件号、零件名称、所属总成
                cell.alignment = left_alignment
            else:
                cell.alignment = data_alignment

        # 行高
        if has_images and item["part_number"] in image_map:
            ws.row_dimensions[row_idx].height = IMAGE_HEIGHT + 10
        else:
            ws.row_dimensions[row_idx].height = 20

    # 插入图片到汇总表
    if has_images:
        seq2 = 0
        for item in leaf_items:
            seq2 += 1
            row_idx = seq2 + 2
            pn = item["part_number"]
            if pn in image_map:
                thumb_path = image_map[pn]
                try:
                    img = XlImage(thumb_path)
                    img.width = IMAGE_WIDTH
                    img.height = IMAGE_HEIGHT
                    ws.add_image(img, f"C{row_idx}")
                except Exception:
                    pass

    col_widths = [30, 8, 18, 25, 40, 10, 18, 14, 14]
    for idx, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    ws.freeze_panes = "A3"
    if leaf_items:
        ws.auto_filter.ref = f"A2:{last_col_letter}{len(leaf_items) + 2}"


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

    # 为叶子零件截图
    image_map = capture_all_images(catia, results, IMAGE_DIR, doc, root_product)

    # 写入 Excel
    root_name = root_product.PartNumber
    root_nom = get_nomenclature(root_product)
    product_display = f"{root_name} - {root_nom}" if root_nom else root_name
    write_to_excel(results, OUTPUT_FILE, image_map, product_name=product_display)
    print()
    print("提取完成！")
    print(f"  - 产品结构表：包含完整树结构层级 + 缩略图")
    print(f"  - 零件汇总表：仅包含叶子零件，便于筛选")
    if image_map:
        print(f"  - 零件图片目录：{IMAGE_DIR}")


if __name__ == "__main__":
    main()
