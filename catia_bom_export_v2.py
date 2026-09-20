# -*- coding: utf-8 -*-
"""CATIA V5 BOM 导出（产品树 + 唯一件汇总 + Create Rough Stock 尺寸）。

使用方法：
1. 在 CATIA 中打开并激活一个 .CATProduct。
2. 运行：python catia_bom_export_v2.py

输出：
* CATIA_BOM_产品名_时间.xlsx
* bom_images\ （每种唯一零件只生成一张图片）
* CATIA_BOM_日志_时间.txt

尺寸口径：对 CATPart 的 MainBody 建立临时 X/Y/Z 极值平面，测量两平面间距。
这与 Create Rough Stock 默认坐标系下的包围盒口径一致。临时几何在每件
测量结束后立即删除，脚本不会保存 CATIA 文档。
"""

from __future__ import annotations

import os
import math
import re
import sys
import time
import traceback
import struct
import tempfile
import uuid
from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

try:
    import pythoncom
    import win32com.client
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as ExcelImage
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from PIL import Image
except ImportError as exc:
    raise SystemExit(f"缺少依赖：{exc}。请运行：pip install pywin32 openpyxl Pillow")


# ----------------------------- 可调配置 ------------------------------------
BASE_DIR = Path(__file__).resolve().parent
EXPORT_ROOT = BASE_DIR / "CATIA_BOM导出"
BOM_OUTPUT_DIR = EXPORT_ROOT / "BOM表"
LOG_DIR = EXPORT_ROOT / "日志"
IMAGE_DIR = EXPORT_ROOT / "零件图片"
# 兼容主导出脚本的现有变量名；Excel 统一写入 BOM表 文件夹。
OUTPUT_DIR = BOM_OUTPUT_DIR
IMAGE_SIZE = (600, 400)  # Excel 中原图的最大显示尺寸；文件本身不缩放
BMP_CAPTURE_FORMAT = 4   # CATIA CatCaptureFormat: catCaptureFormatBMP（无损）
TEMP_PREFIX = "__BOM_TEMP_ROUGH_STOCK__"
# SystemService.Evaluate 的 CATVBScriptLanguage 枚举值。
CAT_VBSCRIPT_LANGUAGE = 1


class BomExportError(RuntimeError):
    """用户可理解的导出错误。"""


@dataclass
class RunLog:
    entries: list[tuple[str, str, str]] = field(default_factory=list)

    def add(self, level: str, message: str, item: str = "") -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.entries.append((level, item, message))
        print(f"[{stamp}] {level:<7} {item} {message}".rstrip())

    def info(self, message: str, item: str = "") -> None:
        self.add("INFO", message, item)

    def warning(self, message: str, item: str = "") -> None:
        self.add("WARNING", message, item)

    def error(self, message: str, item: str = "") -> None:
        self.add("ERROR", message, item)

    def save(self, path: Path) -> None:
        with path.open("w", encoding="utf-8") as handle:
            for level, item, message in self.entries:
                handle.write(f"{level}\t{item}\t{message}\n")


@dataclass
class TreeNode:
    level: int
    tree_number: str
    parent_path: str
    full_path: str
    instance_name: str
    part_number: str
    revision: str
    nomenclature: str
    is_leaf: bool
    product: Any = field(repr=False)
    part_doc: Optional[Any] = field(default=None, repr=False)
    unique_key: str = ""
    tree_quantity: int = 1
    image_path: Optional[Path] = None
    image_status: str = "未生成"


@dataclass
class DimensionResult:
    x: Optional[float] = None
    y: Optional[float] = None
    z: Optional[float] = None
    source: str = "未测量"
    status: str = "未测量"

    @property
    def length(self) -> Optional[float]:
        return self._sorted()[0]

    @property
    def width(self) -> Optional[float]:
        return self._sorted()[1]

    @property
    def height(self) -> Optional[float]:
        return self._sorted()[2]

    def _sorted(self) -> list[Optional[float]]:
        if None in (self.x, self.y, self.z):
            return [None, None, None]
        return sorted([self.x, self.y, self.z], reverse=True)


@dataclass
class MaterialInfo:
    """分别记录产品属性材质和 PartBody 实际材质。"""
    property_material: str = ""
    body_material: str = ""
    comparison: str = "均为空"


@dataclass
class NominalThicknessInfo:
    """零件的名义厚度；不包含几何最小/最大壁厚。"""
    value: str = ""
    source: str = "未识别"
    status: str = "未识别"


@dataclass
class MassResult:
    """CATIA Inertia 计算得到的单件质量及其计算密度。"""
    value_kg: Optional[float] = None
    source: str = "未测量"
    status: str = "未测量"
    density_kg_m3: Optional[float] = None
    density_status: str = "未读取"


@dataclass
class BomItem:
    key: str
    part_number: str
    revision: str
    nomenclature: str
    source_file: str
    representative: TreeNode = field(repr=False)
    occurrences: list[TreeNode] = field(default_factory=list, repr=False)
    dimensions: DimensionResult = field(default_factory=DimensionResult)
    envelope: DimensionResult = field(default_factory=DimensionResult)
    material: MaterialInfo = field(default_factory=MaterialInfo)
    nominal_thickness: NominalThicknessInfo = field(default_factory=NominalThicknessInfo)
    mass: MassResult = field(default_factory=MassResult)
    image_path: Optional[Path] = None
    image_status: str = "未生成"

    @property
    def quantity(self) -> int:
        return len(self.occurrences)


def clean_filename(value: str, fallback: str = "unnamed") -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", value or "").strip(" ._")
    return cleaned or fallback


def com_text(value: Any) -> str:
    """将 CATIA COM 属性安全转为可显示文本。"""
    try:
        return str(value).strip() if value is not None else ""
    except Exception:
        return ""


def connect_catia(log: RunLog) -> Any:
    try:
        catia = win32com.client.GetActiveObject("CATIA.Application")
    except Exception as exc:
        raise BomExportError(
            "未连接到正在运行的 CATIA。请确认 CATIA 已打开 .CATProduct，且 CATIA 与本工具"
            "使用同一 Windows 用户和相同权限（两者都普通运行最简单）。另一台电脑请再检查"
            "pywin32 已安装，并优先使用与 CATIA 位数一致的 Python。"
        ) from exc
    log.info("已连接到 CATIA")
    return catia


def get_active_product_document(catia: Any) -> tuple[Any, Any]:
    try:
        document = catia.ActiveDocument
        if not com_text(document.Name).lower().endswith(".catproduct"):
            raise BomExportError("当前活动文档不是 .CATProduct，请先激活目标产品。")
        return document, document.Product
    except BomExportError:
        raise
    except Exception as exc:
        raise BomExportError("无法读取当前 CATIA 产品文档。") from exc


def get_revision(product: Any) -> str:
    try:
        return com_text(product.Revision)
    except Exception:
        return ""


def get_nomenclature(product: Any) -> str:
    for attribute in ("Nomenclature", "DescriptionInst", "DescriptionRef"):
        try:
            value = com_text(getattr(product, attribute))
            if value:
                return value
        except Exception:
            continue
    return ""


def get_part_document(product: Any, catia: Any) -> Optional[Any]:
    """优先从 ReferenceProduct 取得真实 CATPart 文档，不依赖文件名匹配。"""
    candidates: list[Any] = []
    for candidate_product in (product, _reference_product(product)):
        if candidate_product is None:
            continue
        try:
            candidates.append(candidate_product.Parent)
        except Exception:
            pass

    for document in candidates:
        try:
            if com_text(document.Name).lower().endswith(".catpart") and document.Part:
                return document
        except Exception:
            continue

    # 部分 CATIA 配置下 ReferenceProduct.Parent 不返回 Document，最后扫描已打开文档。
    try:
        for index in range(1, catia.Documents.Count + 1):
            document = catia.Documents.Item(index)
            if not com_text(document.Name).lower().endswith(".catpart"):
                continue
            if com_text(document.Product.PartNumber) == com_text(product.PartNumber):
                return document
    except Exception:
        pass
    return None


def _reference_product(product: Any) -> Optional[Any]:
    try:
        return product.ReferenceProduct
    except Exception:
        return None


def source_file(product: Any, part_doc: Optional[Any]) -> str:
    if part_doc is not None:
        for attribute in ("FullName", "Name"):
            try:
                value = com_text(getattr(part_doc, attribute))
                if value:
                    return value
            except Exception:
                continue
    if product is None:
        return ""
    reference = _reference_product(product)
    if reference is not None:
        try:
            return com_text(reference.PartNumber)
        except Exception:
            pass
    return com_text(product.PartNumber)


def is_leaf(product: Any) -> bool:
    try:
        return product.Products.Count == 0
    except Exception:
        return True


def traverse_product_tree(catia: Any, root: Any, log: RunLog) -> list[TreeNode]:
    nodes: list[TreeNode] = []

    def visit(product: Any, level: int, sequence: str, parent_path: str, tree_quantity: int = 1) -> None:
        part_number = com_text(product.PartNumber) or com_text(product.Name) or "未命名零件"
        leaf = is_leaf(product)
        part_doc = get_part_document(product, catia) if leaf else None
        full_path = f"{parent_path} > {part_number}" if parent_path else part_number
        node = TreeNode(
            level=level,
            tree_number=sequence,
            parent_path=parent_path,
            full_path=full_path,
            instance_name=com_text(product.Name),
            part_number=part_number,
            revision=get_revision(product),
            nomenclature=get_nomenclature(product),
            is_leaf=leaf,
            product=product,
            part_doc=part_doc,
            tree_quantity=tree_quantity,
        )
        nodes.append(node)

        if leaf:
            return
        try:
            children = product.Products
            # 同一父级下的同零件号叶子件只显示一次；数量记录在首个节点。
            # 子总成保持原结构，以免把不同子树错误合并。
            display_children: list[tuple[Any, int]] = []
            seen_leaf_parts: dict[str, int] = {}
            for child_index in range(1, children.Count + 1):
                child = children.Item(child_index)
                if not is_leaf(child):
                    display_children.append((child, 1))
                    continue
                child_number = com_text(child.PartNumber) or com_text(child.Name) or "未命名零件"
                if child_number in seen_leaf_parts:
                    prior_index = seen_leaf_parts[child_number]
                    prior_child, prior_quantity = display_children[prior_index]
                    display_children[prior_index] = (prior_child, prior_quantity + 1)
                else:
                    seen_leaf_parts[child_number] = len(display_children)
                    display_children.append((child, 1))

            for display_index, (child, quantity) in enumerate(display_children, 1):
                visit(child, level + 1, f"{sequence}.{display_index}", full_path, quantity)
        except Exception as exc:
            log.error(f"读取子节点失败：{exc}", full_path)

    visit(root, 0, "1", "")
    return nodes


def create_bom_items(nodes: list[TreeNode]) -> OrderedDict[str, BomItem]:
    """严格按零件号合并：每个零件号只截图一次。"""
    items: OrderedDict[str, BomItem] = OrderedDict()
    for node in nodes:
        if not node.is_leaf:
            continue
        origin = source_file(node.product, node.part_doc)
        key = node.part_number
        node.unique_key = key
        if key not in items:
            items[key] = BomItem(
                key=key,
                part_number=node.part_number,
                revision=node.revision,
                nomenclature=node.nomenclature,
                source_file=origin,
                representative=node,
            )
        # 节点可能代表同一父级下的多个实例，数量必须完整累计。
        items[key].occurrences.extend([node] * node.tree_quantity)
    return items


MATERIAL_PROPERTY_NAMES = ("Material", "MATERIAL", "材料", "材质")
# 仅接受明确表达“名义厚度”的参数，避免将 Pad/Offset 等局部特征的 Thickness
# 错当成零件板厚。产品属性页中的 MaterialThickness/Thickness 也不在此范围内。
NOMINAL_THICKNESS_PARAMETER_NAMES = (
    "NominalThickness",
    "Nominal Thickness",
    "名义厚度",
    "SheetThickness",
    "Sheet Thickness",
    "板厚",
)
SHEET_METAL_THICKNESS_PARAMETER_NAMES = (
    "Sheet Metal Parameters\\Thickness",
    "Sheet Metal Parameters.1\\Thickness",
    "SheetMetalParameters\\Thickness",
    "SheetMetalParameters.1\\Thickness",
    "钣金参数\\厚度",
    "钣金参数.1\\厚度",
)
GEOMETRIC_THICKNESS_MAX_FACES = 120
GEOMETRIC_THICKNESS_MIN_MM = 0.05
GEOMETRIC_THICKNESS_MAX_MM = 100.0
GEOMETRIC_PARALLEL_COSINE = 0.9998  # 面法向夹角不超过约 1.15°


def _read_value(value: Any) -> str:
    """读取 CATIA COM 属性或参数值，兼容属性与零参数方法两种形式。"""
    try:
        if callable(value):
            value = value()
        return com_text(value)
    except Exception:
        return ""


def _read_named_parameter(parameters: Any, names: tuple[str, ...]) -> str:
    if parameters is None:
        return ""
    for name in names:
        try:
            parameter = parameters.Item(name)
            for attribute in ("ValueAsString", "Value"):
                value = _read_value(getattr(parameter, attribute))
                if value:
                    return value
        except Exception:
            continue
    return ""


def _material_name(material: Any) -> str:
    if material is None:
        return ""
    try:
        return _read_value(material.Name)
    except Exception:
        return ""


def _get_material_with_manager(manager: Any, method_name: str, target: Any) -> str:
    """兼容 pywin32 对 CATIA 输出 Material 参数的不同封装方式。"""
    if manager is None or target is None:
        return ""
    try:
        method = getattr(manager, method_name)
    except Exception:
        return ""

    # 在已生成 COM 类型库的环境中，输出参数会直接作为返回对象给出。
    try:
        name = _material_name(method(target))
        if name:
            return name
    except Exception:
        pass

    # 动态 COM 分发需要显式 ByRef IDispatch 包装。
    for variant_type in (
        pythoncom.VT_BYREF | pythoncom.VT_DISPATCH,
        pythoncom.VT_DISPATCH,
    ):
        try:
            material_out = win32com.client.VARIANT(variant_type, None)
            method(target, material_out)
            name = _material_name(getattr(material_out, "value", material_out))
            if name:
                return name
        except Exception:
            continue
    return ""


def _get_property_material(product: Any) -> str:
    """读取 Product 属性页中的 Material，并回退到 Product 级实际材质。"""
    candidates = [product, _reference_product(product)]
    for candidate in candidates:
        if candidate is None:
            continue
        try:
            value = _read_named_parameter(candidate.UserRefProperties, MATERIAL_PROPERTY_NAMES)
            if value:
                return value
        except Exception:
            pass
        try:
            value = _read_named_parameter(candidate.Parameters, MATERIAL_PROPERTY_NAMES)
            if value:
                return value
        except Exception:
            pass
        try:
            manager = candidate.GetItem("CATMatManagerVBExt")
            value = _get_material_with_manager(manager, "GetMaterialOnProduct", candidate)
            if value:
                return value
        except Exception:
            pass
    return ""


def _get_part_body_material(part_doc: Optional[Any]) -> str:
    """读取 MainBody/PartBody 的实际材质及其 Material 参数。"""
    if part_doc is None:
        return ""
    try:
        part = part_doc.Part
        body = part.MainBody
    except Exception:
        return ""

    try:
        manager = part.GetItem("CATMatManagerVBExt")
        value = _get_material_with_manager(manager, "GetMaterialOnBody", body)
        if value:
            return value
    except Exception:
        pass

    # 许多 CATPart 会把图中 PartBody 下的材质以 BodyName\Material 参数公开。
    body_name = com_text(getattr(body, "Name", "")) or "PartBody"
    return _read_named_parameter(
        part.Parameters,
        tuple(f"{body_name}\\{name}" for name in MATERIAL_PROPERTY_NAMES),
    )


def _normalized_material(value: str) -> str:
    return re.sub(r"\s+", "", value.replace("；", ";").replace("，", ",")).casefold()


def read_material_info(product: Any, part_doc: Optional[Any]) -> MaterialInfo:
    property_material = _get_property_material(product)
    body_material = _get_part_body_material(part_doc)
    if property_material and body_material:
        comparison = "一致" if _normalized_material(property_material) == _normalized_material(body_material) else "不一致"
    elif property_material:
        comparison = "仅属性有值"
    elif body_material:
        comparison = "仅PartBody有值"
    else:
        comparison = "均为空"
    return MaterialInfo(property_material, body_material, comparison)


def populate_materials(items: OrderedDict[str, BomItem], log: RunLog) -> None:
    """每个唯一零件号只读取一次材质。"""
    for item in items.values():
        item.material = read_material_info(item.representative.product, item.representative.part_doc)
        if item.material.comparison == "不一致":
            log.warning(
                f"属性 Material 与 PartBody Material 不一致：{item.material.property_material} <> {item.material.body_material}",
                item.part_number,
            )


def _mass_result_from_inertia(inertia: Any, source: str) -> MassResult:
    """从一个 CATIA Inertia 对象读取质量与计算时使用的密度。

    ``Inertia.Density`` 的 1 和 -1 是 CATIA Automation 的状态值，不能把它们
    当作物理密度输出：前者表示未取得材料密度而使用默认密度，后者表示对象含有
    不同密度的几何体。因此质量仍保留，并在表中明确标记密度状态供复核。
    """
    value = float(inertia.Mass)
    if value < 0:
        raise BomExportError(f"CATIA 返回了无效质量：{value}")

    raw_density = float(inertia.Density)
    if abs(raw_density - 1.0) < 1e-9:
        density = None
        density_status = "默认密度（未读取到材料密度）"
    elif abs(raw_density + 1.0) < 1e-9:
        density = None
        density_status = "非均匀密度（多材质/总成）"
    elif raw_density <= 0:
        density = None
        density_status = "未设置或无法读取密度"
    else:
        density = round(raw_density, 3)
        density_status = "已读取（均匀密度）"
    return MassResult(
        value_kg=round(value, 6),
        source=source,
        status="成功",
        density_kg_m3=density,
        density_status=density_status,
    )


def _read_mass_from_product(product: Any, source: str = "CATIA 产品 Inertia（后备）") -> Optional[MassResult]:
    """产品 Inertia 后备读取；Automation 的质量单位为 kg。"""
    if product is None:
        return None
    try:
        return _mass_result_from_inertia(product.GetTechnologicalObject("Inertia"), source)
    except Exception:
        return None


def measure_part_mass(item: BomItem) -> MassResult:
    """以 CATPart 的 PartBody Inertia 为主测量，产品 Inertia 仅在失败时后备。"""
    representative = item.representative
    part_doc = representative.part_doc
    inertia = None
    inertias = None
    partbody_error = ""
    try:
        if part_doc is None:
            raise BomExportError("无法取得 CATPart 文档")
        # 只创建瞬时 Inertia 对象，finally 中会立刻移除，不保存零件。
        part_doc.Activate()
        spa = part_doc.GetWorkbench("SPAWorkbench")
        inertias = spa.Inertias
        inertia = inertias.Add(_find_measurement_body(part_doc.Part))
        return _mass_result_from_inertia(inertia, "PartBody Inertia（主测量）")
    except Exception as exc:
        partbody_error = str(exc)
    finally:
        if inertia is not None and inertias is not None:
            try:
                inertia_name = com_text(inertia.Name)
                if inertia_name:
                    inertias.Remove(inertia_name)
            except Exception:
                pass

    # 某些受保护/轻量化引用无法取得 CATPart 的 PartBody；此时才尝试产品惯性。
    product_candidates = [representative.product, _reference_product(representative.product)]
    if part_doc is not None:
        try:
            product_candidates.append(part_doc.Product)
        except Exception:
            pass
    for product in product_candidates:
        result = _read_mass_from_product(product)
        if result is not None:
            result.source = f"{result.source}；PartBody 失败：{partbody_error}"
            return result
    return MassResult(None, "PartBody Inertia", f"失败：{partbody_error}")


def populate_masses(items: OrderedDict[str, BomItem], product_doc: Optional[Any], log: RunLog) -> None:
    """每个唯一件只测一次 PartBody 质量；结束后恢复总成文档。"""
    for item in items.values():
        log.info("读取质量", item.part_number)
        try:
            item.mass = measure_part_mass(item)
            if item.mass.status == "成功":
                density_text = (
                    f"{item.mass.density_kg_m3:.3f} kg/m³"
                    if item.mass.density_kg_m3 is not None
                    else item.mass.density_status
                )
                log.info(
                    f"质量：{item.mass.value_kg:.6f} kg；密度：{density_text}（{item.mass.source}）",
                    item.part_number,
                )
            else:
                log.warning(item.mass.status, item.part_number)
        finally:
            if product_doc is not None:
                try:
                    product_doc.Activate()
                except Exception as exc:
                    log.warning(f"未能恢复总成文档：{exc}", item.part_number)


def _parameter_name(parameter: Any) -> str:
    try:
        return _read_value(parameter.Name)
    except Exception:
        return ""


def _parameter_text(parameter: Any) -> str:
    """保留 CATIA 参数带单位的显示值，例如 ``1.2mm``。"""
    for attribute in ("ValueAsString", "Value"):
        try:
            value = _read_value(getattr(parameter, attribute))
            if value:
                return value
        except Exception:
            continue
    return ""


def _compact_parameter_name(value: str) -> str:
    return re.sub(r"[\s_.\\\\/-]+", "", value).casefold()


def _thickness_parameter_source(name: str) -> str:
    """判断参数名是否足够明确，可安全作为名义厚度。"""
    compact = _compact_parameter_name(name)
    # 明确排除用户要求不采用的产品属性同名参数。
    if "materialthickness" in compact:
        return ""

    custom_names = {_compact_parameter_name(item) for item in NOMINAL_THICKNESS_PARAMETER_NAMES}
    if compact in custom_names or any(compact.endswith(custom_name) for custom_name in custom_names):
        return "自定义名义厚度参数"

    has_thickness_word = any(word in compact for word in ("thickness", "厚度", "板厚"))
    is_sheet_metal_scope = any(word in compact for word in ("sheetmetal", "钣金", "generativesheetmetal"))
    if has_thickness_word and is_sheet_metal_scope:
        return "CATIA 钣金参数"
    return ""


def _read_nominal_thickness(part_doc: Optional[Any]) -> NominalThicknessInfo:
    """读取钣金或明确命名的名义厚度参数，不进行几何壁厚推算。"""
    if part_doc is None:
        return NominalThicknessInfo(status="未找到 CATPart")
    try:
        parameters = part_doc.Part.Parameters
    except Exception:
        return NominalThicknessInfo(status="无法读取零件参数")

    # 先直接按常见钣金参数路径读取，避免在大量参数中遗漏。
    for name in SHEET_METAL_THICKNESS_PARAMETER_NAMES:
        try:
            parameter = parameters.Item(name)
            value = _parameter_text(parameter)
            if value:
                return NominalThicknessInfo(value, "CATIA 钣金参数", "已识别")
        except Exception:
            continue

    # 不同 CATIA 语言/版本的参数集名称有差异；再遍历所有参数按语义识别。
    try:
        count = parameters.Count
    except Exception:
        return NominalThicknessInfo(status="无法枚举零件参数")
    for index in range(1, count + 1):
        try:
            parameter = parameters.Item(index)
            source = _thickness_parameter_source(_parameter_name(parameter))
            if not source:
                continue
            value = _parameter_text(parameter)
            if value:
                return NominalThicknessInfo(value, source, "已识别")
        except Exception:
            continue
    return NominalThicknessInfo(status="未找到钣金或名义厚度参数")


def populate_nominal_thicknesses(items: OrderedDict[str, BomItem], log: RunLog) -> None:
    """在总成环境中读取每个唯一零件的加厚曲面厚度，不激活 CATPart。"""
    for item in items.values():
        item.nominal_thickness = measure_nominal_thickness(item.representative.part_doc)
        if item.nominal_thickness.status == "已识别":
            log.info(
                f"名义厚度：{item.nominal_thickness.value}（{item.nominal_thickness.source}）",
                item.part_number,
            )
        else:
            log.warning(item.nominal_thickness.status, item.part_number)


def _numeric_components(value: Any, expected_count: int) -> Optional[tuple[float, ...]]:
    """将 pywin32 返回值或 SAFEARRAY 输出转换为数值元组。"""
    if hasattr(value, "value"):
        value = value.value
    if not isinstance(value, (list, tuple)) or len(value) < expected_count:
        return None
    try:
        return tuple(float(value[index]) for index in range(expected_count))
    except (TypeError, ValueError):
        return None


def _catia_script_array(catia: Any, com_object: Any, method_name: str, size: int) -> Optional[tuple[float, ...]]:
    """在 CATIA 进程内调用带数组输出参数的 API。

    pywin32 对 GetPrincipalAxes / GetPoint 一类 ``ByRef SAFEARRAY`` 的封装在
    部分 V5 版本中会丢失输出。SystemService.Evaluate 让 CATIA 自己运行一小段
    VBScript，再把完整数组作为返回值传回 Python。
    """
    if catia is None:
        return None
    script = (
        "Function CATBOM_ReadArray(obj)\n"
        f"  Dim values({size - 1})\n"
        f"  obj.{method_name} values\n"
        "  CATBOM_ReadArray = values\n"
        "End Function"
    )
    try:
        values = catia.SystemService.Evaluate(
            script, CAT_VBSCRIPT_LANGUAGE, "CATBOM_ReadArray", [com_object]
        )
        return _numeric_components(values, size)
    except Exception:
        return None


def _get_plane_components(measurable: Any) -> Optional[tuple[float, ...]]:
    """兼容不同 CATIA/pywin32 版本的 Measurable.GetPlane 输出形式。"""
    try:
        result = measurable.GetPlane()
        components = _numeric_components(result, 9)
        if components:
            return components
    except Exception:
        pass

    # 未生成 COM 类型库时，GetPlane 通常需要显式传入 SAFEARRAY。
    for variant_type in (
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_R8,
    ):
        try:
            output = win32com.client.VARIANT(variant_type, [0.0] * 9)
            result = measurable.GetPlane(output)
            components = _numeric_components(result, 9) or _numeric_components(output, 9)
            if components:
                return components
        except Exception:
            continue

    # 个别环境直接修改传入的 Python list。
    try:
        output_list = [0.0] * 9
        result = measurable.GetPlane(output_list)
        return _numeric_components(result, 9) or _numeric_components(output_list, 9)
    except Exception:
        return None


def _cross_product(first: tuple[float, float, float], second: tuple[float, float, float]) -> tuple[float, float, float]:
    return (
        first[1] * second[2] - first[2] * second[1],
        first[2] * second[0] - first[0] * second[2],
        first[0] * second[1] - first[1] * second[0],
    )


def _normalize_vector(vector: tuple[float, float, float]) -> Optional[tuple[float, float, float]]:
    length = math.sqrt(sum(value * value for value in vector))
    if length <= 1e-12:
        return None
    return tuple(value / length for value in vector)


def _dot_product(first: tuple[float, float, float], second: tuple[float, float, float]) -> float:
    return sum(first[index] * second[index] for index in range(3))


def _selection_count(selection: Any) -> int:
    for attribute in ("Count2", "Count"):
        try:
            return int(getattr(selection, attribute))
        except Exception:
            continue
    return 0


def _search_topological_faces(selection: Any) -> int:
    """使用与 CATIA 会话语言无关的面搜索语法，并兼容旧版本分隔符。"""
    # CATIA Search Language 的正式写法为 .Face,all。Topology.Face 在部分
    # V5 配置中不是有效 TransFormat 查询，调用 Search 会直接抛异常。
    queries = (
        ".Face,all",
        ".Face;all",
        "Topology.Face,all",
        "Topology.Face;all",
        "CATPrtSearch.Face,all",
        "CATPrtSearch.Face;all",
    )
    failures: list[str] = []
    query_succeeded = False
    for query in queries:
        try:
            selection.Clear()
            selection.Search(query)
            query_succeeded = True
            count = _selection_count(selection)
            if count:
                return count
        except Exception as exc:
            failures.append(f"{query}: {exc}")
    if query_succeeded:
        return 0
    raise BomExportError("CATIA 无法搜索拓扑面：" + "；".join(failures))


def _selected_face_reference(part: Any, selected_element: Any) -> tuple[Optional[Any], Optional[Any]]:
    """取得所选面的 Reference 及原对象，兼容 CATIA 版本差异。"""
    try:
        face_object = selected_element.Value
    except Exception:
        return None, None
    try:
        return part.CreateReferenceFromObject(face_object), face_object
    except Exception:
        try:
            return selected_element.Reference, face_object
        except Exception:
            return None, face_object


def _measured_face_distance(measurable: Any, second_reference: Any, second_object: Any) -> Optional[float]:
    """CATIA 示例在版本间对第二参数接受 Reference 或对象有所不同。"""
    for candidate in (second_reference, second_object):
        if candidate is None:
            continue
        try:
            distance = abs(float(measurable.GetMinimumDistance(candidate)))
            if math.isfinite(distance):
                return distance
        except Exception:
            continue
    return None


def _com_interface_name(value: Any) -> str:
    """尽量取得 COM 对象的接口名，例如 ThickSurface。"""
    try:
        type_info = value._oleobj_.GetTypeInfo()
        return com_text(type_info.GetDocumentation(-1)[0])
    except Exception:
        return type(value).__name__


def _length_mm(length_parameter: Any) -> Optional[float]:
    """读取 CATIA Length 的数值并统一转换为 mm。"""
    try:
        raw_value = _read_value(getattr(length_parameter, "Value"))
        if raw_value:
            return float(raw_value)
    except (TypeError, ValueError):
        pass
    except Exception:
        pass

    # 个别 CATIA 配置只提供带单位的文本值。
    try:
        raw_text = _read_value(getattr(length_parameter, "ValueAsString"))
        match = re.search(r"[-+]?\d+(?:\.\d+)?", raw_text.replace(",", "."))
        if not match:
            return None
        value = float(match.group())
        unit = raw_text[match.end():].strip().casefold()
        for suffix, factor in (("mm", 1.0), ("cm", 10.0), ("m", 1000.0), ("in", 25.4)):
            if unit.startswith(suffix):
                return value * factor
        return value  # CATIA 长度显示未带单位时，按当前零件单位（通常为 mm）处理。
    except Exception:
        return None


def _read_thick_surface_nominal_thickness(part_doc: Optional[Any]) -> NominalThicknessInfo:
    """从 PartBody 中的 ThickSurface/加厚曲面特征读取实体生成厚度。"""
    if part_doc is None:
        return NominalThicknessInfo(status="未找到 CATPart")
    try:
        bodies = part_doc.Part.Bodies
    except Exception:
        return NominalThicknessInfo(status="无法读取 CATPart 的 Body 特征")

    candidates: list[tuple[str, float]] = []
    read_failures: list[str] = []
    try:
        body_count = bodies.Count
    except Exception:
        body_count = 0
    for body_index in range(1, body_count + 1):
        try:
            shapes = bodies.Item(body_index).Shapes
            shape_count = shapes.Count
        except Exception:
            continue
        for shape_index in range(1, shape_count + 1):
            try:
                shape = shapes.Item(shape_index)
                feature_name = com_text(getattr(shape, "Name", ""))
                feature_type = _com_interface_name(shape)
                identity = f"{feature_name} {feature_type}".casefold()
                if not any(token in identity for token in ("thicksurface", "加厚曲面", "厚曲面")):
                    continue
                top = _length_mm(getattr(shape, "TopOffset"))
                bottom = _length_mm(getattr(shape, "BotOffset"))
                if top is None or bottom is None:
                    read_failures.append(feature_name or feature_type or "加厚曲面")
                    continue
                thickness = abs(top) + abs(bottom)
                if thickness > 0:
                    candidates.append((feature_name or feature_type or "加厚曲面", thickness))
            except Exception:
                continue

    if not candidates:
        if read_failures:
            return NominalThicknessInfo(status=f"找到加厚曲面但无法读取偏移值：{', '.join(read_failures)}")
        return NominalThicknessInfo(status="未找到加厚曲面特征")

    # 多个加厚曲面只有在厚度一致时才自动采用，避免把局部加强厚度误写为板厚。
    unique_values = sorted({round(value, 3) for _, value in candidates})
    if len(unique_values) != 1:
        names = ", ".join(name for name, _ in candidates)
        values = ", ".join(f"{value:.3f}" for value in unique_values)
        return NominalThicknessInfo(status=f"多个加厚曲面厚度不一致（{values} mm）：{names}")

    thickness = unique_values[0]
    return NominalThicknessInfo(
        value=f"{thickness:.3f} mm",
        source=f"加厚曲面特征（{len(candidates)} 个）",
        status="已识别",
    )


def measure_nominal_thickness_from_faces(part_doc: Optional[Any]) -> NominalThicknessInfo:
    """从实体几何实际测量名义厚度。

    方法：提取平面面，筛选法向相反的平行面；只有其 CATIA 最小距离与两
    平面真实间距一致时，才将该距离视为有效壁厚候选。最后按匹配面面积和
    出现次数选择最具代表性的候选值。该过程不读取 Thickness 属性或参数。
    """
    if part_doc is None:
        return NominalThicknessInfo(status="未找到 CATPart，无法几何测量")

    selection = None
    try:
        part_doc.Activate()
        part = part_doc.Part
        spa = part_doc.GetWorkbench("SPAWorkbench")
        selection = part_doc.Selection
        # 搜索只读的拓扑面；所有筛选都在 Python 内完成。
        selected_count = _search_topological_faces(selection)
        if selected_count < 2:
            return NominalThicknessInfo(status="实体平面不足，无法几何测量")

        faces: list[dict[str, Any]] = []
        for index in range(1, selected_count + 1):
            try:
                selected = selection.Item2(index) if hasattr(selection, "Item2") else selection.Item(index)
                reference, face_object = _selected_face_reference(part, selected)
                if reference is None:
                    continue
                measurable = spa.GetMeasurable(reference)
                plane = _get_plane_components(measurable)
                if plane is None:
                    continue  # 圆柱面、圆锥面、自由曲面不是本算法的可靠输入。
                normal = _normalize_vector(
                    _cross_product((plane[3], plane[4], plane[5]), (plane[6], plane[7], plane[8]))
                )
                if normal is None:
                    continue
                try:
                    area = max(float(measurable.Area), 0.0)
                except Exception:
                    area = 0.0
                faces.append({
                    "reference": reference,
                    "object": face_object,
                    "measurable": measurable,
                    "origin": (plane[0], plane[1], plane[2]),
                    "normal": normal,
                    "area": area,
                })
            except Exception:
                continue

        if len(faces) < 2:
            return NominalThicknessInfo(status="未找到可测量的平面面")
        # 特征非常多时优先保留大面，避免 O(n²) 测量让 CATIA 长时间无响应。
        faces.sort(key=lambda face: face["area"], reverse=True)
        faces = faces[:GEOMETRIC_THICKNESS_MAX_FACES]

        clusters: dict[float, dict[str, float]] = {}
        for first_index, first in enumerate(faces[:-1]):
            for second in faces[first_index + 1:]:
                # 板材两侧面的法向应相反且几乎共线。
                if _dot_product(first["normal"], second["normal"]) > -GEOMETRIC_PARALLEL_COSINE:
                    continue
                delta = tuple(second["origin"][axis] - first["origin"][axis] for axis in range(3))
                plane_spacing = abs(_dot_product(delta, first["normal"]))
                if not GEOMETRIC_THICKNESS_MIN_MM <= plane_spacing <= GEOMETRIC_THICKNESS_MAX_MM:
                    continue
                actual_distance = _measured_face_distance(first["measurable"], second["reference"], second["object"])
                if actual_distance is None:
                    continue
                # 最短连线应垂直于这对平面；否则两面没有对应的投影区域，不能代表壁厚。
                tolerance = max(0.02, plane_spacing * 0.01)
                if abs(actual_distance - plane_spacing) > tolerance:
                    continue
                key = round(plane_spacing, 2)
                cluster = clusters.setdefault(key, {"area": 0.0, "count": 0.0})
                cluster["area"] += min(first["area"], second["area"])
                cluster["count"] += 1.0

        if not clusters:
            return NominalThicknessInfo(status="没有找到可确认的平行面壁厚")
        # 实体中最常出现、对应面积最大的真实面间距即为名义厚度。
        thickness, evidence = max(
            clusters.items(),
            key=lambda candidate: (candidate[1]["area"] * candidate[1]["count"], candidate[1]["count"]),
        )
        return NominalThicknessInfo(
            value=f"{thickness:.2f} mm",
            source=f"几何实测（{int(evidence['count'])} 对平行面）",
            status="已识别",
        )
    except Exception as exc:
        return NominalThicknessInfo(status=f"几何测量失败：{exc}")
    finally:
        if selection is not None:
            try:
                selection.Clear()
            except Exception:
                pass


def measure_nominal_thickness(part_doc: Optional[Any]) -> NominalThicknessInfo:
    """只读取已有加厚曲面特征；不激活零件、不切换工作台、不做面搜索。"""
    return _read_thick_surface_nominal_thickness(part_doc)


def _append_shape(container: Any, shape: Any) -> Any:
    container.AppendHybridShape(shape)
    return shape


def _delete_temp_body(part_doc: Any, temp_body: Any) -> None:
    """用 CATIA 直接删除 API 清理临时几何集。

    不使用 ``Selection.Delete``：部分 CATIA V5/PLM 环境会把该调用误判为
    “复制数据”，弹出“选定的元素不可用于此操作”。
    """
    part_doc.Activate()
    part = part_doc.Part
    try:
        part.InWorkObject = part.MainBody
    except Exception:
        pass
    reference = part.CreateReferenceFromObject(temp_body)
    part.HybridShapeFactory.DeleteObjectForDatum(reference)


def _find_measurement_body(part: Any) -> Any:
    """优先 PartBody；没有实体时再寻找含 Shape 的其他 Body。"""
    candidates = []
    try:
        candidates.append(part.MainBody)
    except Exception:
        pass
    try:
        for index in range(1, part.Bodies.Count + 1):
            candidates.append(part.Bodies.Item(index))
    except Exception:
        pass
    seen = set()
    for body in candidates:
        try:
            identifier = com_text(body.Name)
            if identifier in seen:
                continue
            seen.add(identifier)
            if body.Shapes.Count > 0:
                return body
        except Exception:
            continue
    raise BomExportError("所有 Body 都没有可测量实体")


def _remove_stale_temp_bodies(part_doc: Any) -> None:
    """清理异常中断的旧测量几何，避免污染下一次尺寸计算。"""
    part_doc.Activate()
    hybrid_bodies = part_doc.Part.HybridBodies
    stale_bodies = []
    for index in range(1, hybrid_bodies.Count + 1):
        body = hybrid_bodies.Item(index)
        if com_text(body.Name).startswith(TEMP_PREFIX):
            stale_bodies.append(body)
    failures = []
    for body in reversed(stale_bodies):
        try:
            _delete_temp_body(part_doc, body)
        except Exception as exc:
            failures.append(str(exc))
    if failures:
        raise BomExportError("无法清理上次残留的临时测量几何：" + "；".join(failures))


def _compute(shape: Any) -> None:
    """逐个计算临时特征，避免整件 Part.Update 因无关失效特征而失败。"""
    shape.Compute()


def _get_point_coordinates(measurable: Any, catia: Any = None) -> Optional[tuple[float, float, float]]:
    """兼容不同 CATIA/pywin32 版本的 Measurable.GetPoint 输出形式。"""
    coordinates = _catia_script_array(catia, measurable, "GetPoint", 3)
    if coordinates is not None:
        return coordinates
    try:
        coordinates = _numeric_components(measurable.GetPoint(), 3)
        if coordinates:
            return coordinates
    except Exception:
        pass
    for variant_type in (
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_R8,
        pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
    ):
        try:
            output = win32com.client.VARIANT(variant_type, [0.0] * 3)
            result = measurable.GetPoint(output)
            coordinates = _numeric_components(result, 3) or _numeric_components(output, 3)
            if coordinates:
                return coordinates
        except Exception:
            continue
    try:
        output_list = [0.0] * 3
        result = measurable.GetPoint(output_list)
        return _numeric_components(result, 3) or _numeric_components(output_list, 3)
    except Exception:
        return None


def _get_explicit_point_coordinates(
    part_doc: Any, point: Any, allow_local_fallback: bool = False
) -> Optional[tuple[float, float, float]]:
    """读取显式 PointCoord 的全局坐标。

    ``PointCoordWithReference(0,0,0, Extremum)`` 的 GetCoordinates 返回的是
    *相对 Extremum 的偏移*，即恒为 0；因此必须优先走 SPA Measurable.GetPoint。
    """
    try:
        reference = part_doc.Part.CreateReferenceFromObject(point)
        coordinates = _get_point_coordinates(
            part_doc.GetWorkbench("SPAWorkbench").GetMeasurable(reference),
            part_doc.Application,
        )
        if coordinates is not None:
            return coordinates
    except Exception:
        pass
    # 仅保留对不带参考的独立 PointCoord 的兼容回退；测量路径不会使用其零偏移值。
    if not allow_local_fallback:
        return None
    for method_name in ("GetCoordinates",):
        method = getattr(point, method_name, None)
        if method is None:
            continue
        try:
            coordinates = _numeric_components(method(), 3)
            if coordinates:
                return coordinates
        except Exception:
            pass
        for variant_type in (
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_R8,
            pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
            pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
        ):
            try:
                output = win32com.client.VARIANT(variant_type, [0.0] * 3)
                result = method(output)
                coordinates = _numeric_components(result, 3) or _numeric_components(output, 3)
                if coordinates:
                    return coordinates
            except Exception:
                continue
        try:
            output_list = [0.0] * 3
            result = method(output_list)
            coordinates = _numeric_components(result, 3) or _numeric_components(output_list, 3)
            if coordinates:
                return coordinates
        except Exception:
            pass
    return None


def _extract_measurement_body_surface(part_doc: Any, part: Any, temp_body: Any, factory: Any, body: Any) -> Any:
    """提取 PartBody 的一个面并传播到整个实体，供 Extremum 稳定计算。

    V5 的 HybridShapeExtremum 对实体 Body 的支持在不同版本中不一致；而对
    “从该 Body 面传播出的 Extract”是稳定的。这里仅使用 Selection.Search，
    不会复制、粘贴或删除用户几何。
    """
    selection = part_doc.Selection
    selected_values: list[Any] = []
    try:
        selection.Clear()
        selection.Add(body)
        for query in ("Topology.Face,sel", "Topology.CGMFace,sel"):
            try:
                selection.Search(query)
                if selection.Count:
                    for index in range(1, selection.Count + 1):
                        try:
                            item = selection.Item2(index)
                        except Exception:
                            item = selection.Item(index)
                        for attribute in ("Reference", "Value"):
                            try:
                                candidate = getattr(item, attribute)
                                if candidate is not None:
                                    selected_values.append(candidate)
                            except Exception:
                                continue
                        if selected_values:
                            break
                if selected_values:
                    break
            except Exception:
                continue
    finally:
        try:
            selection.Clear()
        except Exception:
            pass

    for face in selected_values:
        try:
            extract = factory.AddNewExtract(face)
            try:
                extract.PropagationType = 1  # 连通传播，覆盖整个 PartBody。
            except Exception:
                pass
            try:
                extract.ComplementaryExtract = False
                extract.IsFederated = False
            except Exception:
                pass
            extract = _append_shape(temp_body, extract)
            _update_measurement_shape(part, extract)
            return part.CreateReferenceFromObject(extract)
        except Exception:
            continue
    raise BomExportError("无法从 PartBody 取得可用于 Rough Stock 的实体面")


def _update_measurement_shape(part: Any, shape: Any) -> None:
    """优先仅更新本次特征，避免整件 Update；旧版本回退到 Compute。"""
    try:
        part.UpdateObject(shape)
    except Exception:
        _compute(shape)


def measure_rough_stock_dimensions(
    part_doc: Any,
    directions: Optional[tuple[tuple[float, float, float], ...]] = None,
    source: str = "Create Rough Stock 规则（零件本地 XYZ）",
) -> DimensionResult:
    """返回指定轴系下的 Rough Stock 包络尺寸（mm）。

    CATIA Automation 没有直接暴露 Create Rough Stock 对话框中的 six bounds。
    这里对 MainBody 创建 X/Y/Z 的极值点，直接读取点坐标并计算 max-min。
    该口径与 Rough Stock 对话框中的 DX/DY/DZ 一致，不使用不稳定的平面间距测量。
    """
    part = part_doc.Part
    temp_body = None
    result = DimensionResult(source=source)
    try:
        # Selection.Delete 和 CATIA 几何工厂均要求被操作的 CATPart 是活动文档。
        part_doc.Activate()
        _remove_stale_temp_bodies(part_doc)
        measurement_body = _find_measurement_body(part)

        hybrid_bodies = part.HybridBodies
        temp_body = hybrid_bodies.Add()
        temp_body.Name = f"{TEMP_PREFIX}{uuid.uuid4().hex}"
        part.InWorkObject = temp_body
        factory = part.HybridShapeFactory

        body_ref = _extract_measurement_body_surface(
            part_doc, part, temp_body, factory, measurement_body
        )
        extrema_by_axis: list[tuple[tuple[float, float, float], Any, Any]] = []

        active_directions = directions or ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0))
        if len(active_directions) != 3:
            raise BomExportError("Rough Stock 轴系必须包含三个方向")
        for vector in active_directions:
            normalized_vector = _normalize_vector(vector)
            if normalized_vector is None:
                raise BomExportError("Rough Stock 轴系包含无效方向")
            direction = _append_shape(temp_body, factory.AddNewDirectionByCoord(*normalized_vector))
            _compute(direction)
            minimum = _append_shape(temp_body, factory.AddNewExtremum(body_ref, direction, 0))
            maximum = _append_shape(temp_body, factory.AddNewExtremum(body_ref, direction, 1))
            _update_measurement_shape(part, minimum)
            _update_measurement_shape(part, maximum)
            # Extremum 是关联特征而非可测量 Point。创建零偏移 PointCoord 后，
            # 才能稳定通过 GetCoordinates / SPA.GetPoint 取得坐标。
            minimum_point = _append_shape(
                temp_body,
                factory.AddNewPointCoordWithReference(
                    0.0, 0.0, 0.0, part.CreateReferenceFromObject(minimum)
                ),
            )
            maximum_point = _append_shape(
                temp_body,
                factory.AddNewPointCoordWithReference(
                    0.0, 0.0, 0.0, part.CreateReferenceFromObject(maximum)
                ),
            )
            _update_measurement_shape(part, minimum_point)
            _update_measurement_shape(part, maximum_point)
            extrema_by_axis.append((normalized_vector, minimum_point, maximum_point))

        distances: list[float] = []
        for axis_vector, minimum_point, maximum_point in extrema_by_axis:
            min_point = _get_explicit_point_coordinates(part_doc, minimum_point)
            max_point = _get_explicit_point_coordinates(part_doc, maximum_point)
            if min_point is None or max_point is None:
                raise BomExportError("无法读取 Rough Stock 极值点坐标")
            delta = tuple(max_point[index] - min_point[index] for index in range(3))
            distances.append(round(abs(_dot_product(delta, axis_vector)), 3))

        if any(value <= 0 for value in distances):
            raise BomExportError(f"得到无效尺寸 {distances} mm")

        result = DimensionResult(
            x=distances[0],
            y=distances[1],
            z=distances[2],
            source=source,
            status="成功",
        )
    except Exception as exc:
        result = DimensionResult(
            source=source,
            status=f"失败：{exc}",
        )
    finally:
        if temp_body is not None:
            try:
                _delete_temp_body(part_doc, temp_body)
            except Exception as cleanup_exc:
                # 尺寸已经成功时，不能因清理告警而让 Excel 丢弃有效测量值。
                if result.status == "成功":
                    result.source = f"{result.source}；清理警告：{cleanup_exc}"
                else:
                    result.status = f"{result.status}；警告：临时测量几何未完全清理：{cleanup_exc}"
    return result


def populate_rough_stock_dimensions(
    items: OrderedDict[str, BomItem],
    product_doc: Optional[Any],
    log: RunLog,
) -> None:
    """读取每个唯一件的包装包络尺寸，并在每件后恢复总成。"""
    for item in items.values():
        log.info("测量包装包络（PartBody 惯性主轴 Rough Stock）", item.part_number)
        try:
            if item.representative.part_doc is None:
                item.dimensions = DimensionResult(
                    source="Create Rough Stock 规则（零件本地 XYZ）",
                    status="失败：无法取得 CATPart 文档",
                )
            else:
                item.dimensions = measure_packaging_envelope(item.representative.part_doc)
            if item.dimensions.status == "成功":
                log.info(
                    f"包络：{item.dimensions.length:.3f} × {item.dimensions.width:.3f} × {item.dimensions.height:.3f} mm",
                    item.part_number,
                )
            else:
                log.warning(item.dimensions.status, item.part_number)
        finally:
            if product_doc is not None:
                try:
                    product_doc.Activate()
                except Exception as exc:
                    log.warning(f"未能恢复总成文档：{exc}", item.part_number)


def _read_principal_axes(inertia: Any, catia: Any = None) -> Optional[tuple[tuple[float, float, float], ...]]:
    """读取 CATIA Inertia 的三个主轴，返回正交单位向量。"""
    components = _catia_script_array(catia, inertia, "GetPrincipalAxes", 9)
    try:
        if components is None:
            components = _numeric_components(inertia.GetPrincipalAxes(), 9)
    except Exception:
        pass
    if components is None:
        for variant_type in (
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_R8,
            pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
            pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
        ):
            try:
                output = win32com.client.VARIANT(variant_type, [0.0] * 9)
                result = inertia.GetPrincipalAxes(output)
                components = _numeric_components(result, 9) or _numeric_components(output, 9)
                if components is not None:
                    break
            except Exception:
                continue
    if components is None:
        try:
            output_list = [0.0] * 9
            result = inertia.GetPrincipalAxes(output_list)
            components = _numeric_components(result, 9) or _numeric_components(output_list, 9)
        except Exception:
            pass
    if components is None:
        return None

    # CATIA 返回数组按列存放：A1=(0,3,6)，A2=(1,4,7)，A3=(2,5,8)。
    first = _normalize_vector((components[0], components[3], components[6]))
    second = _normalize_vector((components[1], components[4], components[7]))
    third = _normalize_vector((components[2], components[5], components[8]))
    if first is None or second is None or third is None:
        return None
    # 消除数值误差，并确保三轴构成右手坐标系。
    second = _normalize_vector(_cross_product(third, first))
    if second is None:
        return None
    third = _normalize_vector(_cross_product(first, second))
    if third is None:
        return None
    return first, second, third


def _get_inertia_axes(part_doc: Any) -> Optional[tuple[tuple[float, float, float], ...]]:
    """以 MainBody/PartBody 创建瞬时 Inertia 对象取得主轴，结束后立即移除。"""
    inertia = None
    inertias = None
    try:
        spa = part_doc.GetWorkbench("SPAWorkbench")
        inertias = spa.Inertias
        inertia = inertias.Add(_find_measurement_body(part_doc.Part))
        return _read_principal_axes(inertia, part_doc.Application)
    except Exception:
        return None
    finally:
        if inertia is not None and inertias is not None:
            try:
                inertia_name = com_text(inertia.Name)
                if inertia_name:
                    inertias.Remove(inertia_name)
            except Exception:
                pass


def measure_inertia_rough_stock_dimensions(part_doc: Any) -> DimensionResult:
    """以 PartBody 的惯性主轴作为 Rough Stock 轴系，得到更紧凑的包装包络。"""
    source = "Create Rough Stock（PartBody 惯性主轴）"
    try:
        part_doc.Activate()
        inertia_axes = _get_inertia_axes(part_doc)
        if inertia_axes is None:
            return DimensionResult(source=source, status="失败：无法读取 PartBody 惯性主轴")
        return measure_rough_stock_dimensions(part_doc, inertia_axes, source)
    except Exception as exc:
        return DimensionResult(source=source, status=f"失败：{exc}")


def measure_packaging_envelope(part_doc: Any) -> DimensionResult:
    """按可靠性依次尝试 CATIA Rough Stock、STL 网格主方向与三视图包络。"""
    inertia_result = measure_inertia_rough_stock_dimensions(part_doc)
    if inertia_result.status == "成功":
        return inertia_result

    projection_source = "后备：PartBody 三视图投影外接矩形（零件 XYZ）"
    projection_result = measure_rough_stock_dimensions(part_doc, source=projection_source)
    if projection_result.status == "成功":
        return projection_result

    return DimensionResult(
        source="包装包络",
        status=f"失败：惯性轴 Rough Stock：{inertia_result.status}；三视图投影：{projection_result.status}",
    )


def _iter_stl_vertices(stl_path: Path):
    """逐个产出 STL 顶点，兼容 CATIA 导出的二进制或 ASCII STL。"""
    data = stl_path.read_bytes()
    if len(data) >= 84:
        triangle_count = struct.unpack_from("<I", data, 80)[0]
        expected_size = 84 + triangle_count * 50
        if expected_size == len(data):
            for record in struct.iter_unpack("<12fH", data[84:]):
                yield (float(record[3]), float(record[4]), float(record[5]))
                yield (float(record[6]), float(record[7]), float(record[8]))
                yield (float(record[9]), float(record[10]), float(record[11]))
            return

    vertex_pattern = re.compile(
        rb"\bvertex\s+([-+0-9.eE]+)\s+([-+0-9.eE]+)\s+([-+0-9.eE]+)"
    )
    for match in vertex_pattern.finditer(data):
        try:
            yield tuple(float(value) for value in match.groups())
        except ValueError:
            continue


def _mesh_axes_from_vertices(
    vertices: list[tuple[float, float, float]],
) -> Optional[tuple[tuple[float, float, float], ...]]:
    """用网格顶点 PCA 求近似主轴；无需 NumPy，适合批量导出环境。"""
    if len(vertices) < 4:
        return None
    count = float(len(vertices))
    center = tuple(sum(vertex[index] for vertex in vertices) / count for index in range(3))
    matrix = [[0.0, 0.0, 0.0] for _ in range(3)]
    for vertex in vertices:
        delta = tuple(vertex[index] - center[index] for index in range(3))
        for row in range(3):
            for column in range(3):
                matrix[row][column] += delta[row] * delta[column]

    def multiply(vector: tuple[float, float, float]) -> tuple[float, float, float]:
        return tuple(sum(matrix[row][column] * vector[column] for column in range(3)) for row in range(3))

    def principal(seed: tuple[float, float, float], excluded: list[tuple[float, float, float]]):
        vector = _normalize_vector(seed)
        if vector is None:
            return None
        for _ in range(80):
            candidate = multiply(vector)
            for previous in excluded:
                scalar = _dot_product(candidate, previous)
                candidate = tuple(candidate[index] - scalar * previous[index] for index in range(3))
            updated = _normalize_vector(candidate)
            if updated is None:
                return None
            if abs(abs(_dot_product(updated, vector)) - 1.0) < 1e-10:
                return updated
            vector = updated
        return vector

    first = principal((1.0, 0.37, 0.19), [])
    second = principal((0.23, 1.0, 0.41), [first] if first else [])
    if first is None or second is None:
        return None
    third = _normalize_vector(_cross_product(first, second))
    if third is None:
        return None
    second = _normalize_vector(_cross_product(third, first))
    if second is None:
        return None
    return first, second, third


def _dimensions_from_mesh(
    stl_path: Path,
    axes: tuple[tuple[float, float, float], ...],
) -> Optional[tuple[float, float, float]]:
    minimum = [float("inf")] * 3
    maximum = [float("-inf")] * 3
    found = False
    for vertex in _iter_stl_vertices(stl_path):
        found = True
        for index, axis in enumerate(axes):
            position = _dot_product(vertex, axis)
            minimum[index] = min(minimum[index], position)
            maximum[index] = max(maximum[index], position)
    if not found:
        return None
    dimensions = tuple(round(maximum[index] - minimum[index], 3) for index in range(3))
    return dimensions if all(value > 0 for value in dimensions) else None


def measure_stl_packaging_envelope(part_doc: Any) -> DimensionResult:
    """临时导出 STL，求网格主方向包络；绝不创建或保存 CATIA 特征。"""
    try:
        part_doc.Activate()
        with tempfile.TemporaryDirectory(prefix="catia_bom_mesh_") as directory:
            stl_path = Path(directory) / f"{uuid.uuid4().hex}.stl"
            part_doc.ExportData(str(stl_path), "stl")
            if not stl_path.exists() or stl_path.stat().st_size < 84:
                raise BomExportError("CATIA 未生成 STL 临时网格文件")

            # 对大网格进行均匀降采样以求轴向；最终尺寸仍遍历全部顶点，保证边界准确。
            sample: list[tuple[float, float, float]] = []
            stride = 1
            for index, vertex in enumerate(_iter_stl_vertices(stl_path)):
                if index % stride == 0:
                    sample.append(vertex)
                if len(sample) >= 120000:
                    sample = sample[::2]
                    stride *= 2
            local_axes = ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0))
            mesh_axes = _mesh_axes_from_vertices(sample)
            if mesh_axes is not None:
                dimensions = _dimensions_from_mesh(stl_path, mesh_axes)
                if dimensions is not None:
                    return DimensionResult(
                        x=dimensions[0], y=dimensions[1], z=dimensions[2],
                        source="后备：STL 网格主方向近似包装包络",
                        status="成功",
                    )
            dimensions = _dimensions_from_mesh(stl_path, local_axes)
            if dimensions is None:
                raise BomExportError("STL 网格没有可用的三维顶点")
            return DimensionResult(
                x=dimensions[0], y=dimensions[1], z=dimensions[2],
                source="后备：STL 网格三视图投影外接尺寸（零件 XYZ）",
                status="成功",
            )
    except Exception as exc:
        return DimensionResult(source="STL 网格包装包络", status=f"失败：{exc}")


def _measure_envelope_in_frame(
    part_doc: Any,
    directions: tuple[tuple[float, float, float], ...],
    source: str,
) -> DimensionResult:
    """沿一组正交方向计算实体投影尺寸；临时几何一定会清理。"""
    temp_body = None
    result = DimensionResult(source=source, status="未测量")
    try:
        part_doc.Activate()
        _remove_stale_temp_bodies(part_doc)
        part = part_doc.Part
        measurement_body = _find_measurement_body(part)
        temp_body = part.HybridBodies.Add()
        temp_body.Name = f"{TEMP_PREFIX}ENVELOPE_{uuid.uuid4().hex}"
        part.InWorkObject = temp_body
        factory = part.HybridShapeFactory
        origin = _append_shape(temp_body, factory.AddNewPointCoord(0.0, 0.0, 0.0))
        _compute(origin)
        body_ref = part.CreateReferenceFromObject(measurement_body)
        planes: list[tuple[Any, Any]] = []
        for direction_vector in directions:
            direction = _append_shape(temp_body, factory.AddNewDirectionByCoord(*direction_vector))
            _compute(direction)
            axis_line = _append_shape(
                temp_body,
                factory.AddNewLinePtDir(origin, direction, 1000000.0, -1000000.0, True),
            )
            _compute(axis_line)
            minimum = _append_shape(temp_body, factory.AddNewExtremum(body_ref, direction, 0))
            maximum = _append_shape(temp_body, factory.AddNewExtremum(body_ref, direction, 1))
            _compute(minimum)
            _compute(maximum)
            min_plane = _append_shape(temp_body, factory.AddNewPlaneNormal(axis_line, minimum))
            max_plane = _append_shape(temp_body, factory.AddNewPlaneNormal(axis_line, maximum))
            _compute(min_plane)
            _compute(max_plane)
            planes.append((min_plane, max_plane))

        spa = part_doc.GetWorkbench("SPAWorkbench")
        values: list[float] = []
        for min_plane, max_plane in planes:
            measurable = spa.GetMeasurable(part.CreateReferenceFromObject(min_plane))
            values.append(round(abs(float(measurable.GetMinimumDistance(max_plane))), 3))
        if len(values) != 3 or any(value <= 0 for value in values):
            raise BomExportError(f"得到无效包络尺寸 {values} mm")
        result = DimensionResult(x=values[0], y=values[1], z=values[2], source=source, status="成功")
    except Exception as exc:
        result = DimensionResult(source=source, status=f"失败：{exc}")
    finally:
        if temp_body is not None:
            try:
                _delete_temp_body(part_doc, temp_body)
            except Exception as cleanup_exc:
                result.source = f"{result.source}；清理警告：{cleanup_exc}"
    return result


def measure_approximate_min_envelope(part_doc: Optional[Any]) -> DimensionResult:
    """以零件 XYZ 和惯性主轴两组方向近似最小包络，返回体积更小的一组。"""
    if part_doc is None:
        return DimensionResult(source="近似最小包络", status="失败：无法取得 CATPart 文档")
    local_axes = ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0))
    candidates = [_measure_envelope_in_frame(part_doc, local_axes, "近似最小包络（零件 XYZ）")]
    inertia_axes = _get_inertia_axes(part_doc)
    if inertia_axes is not None:
        candidates.append(_measure_envelope_in_frame(part_doc, inertia_axes, "近似最小包络（惯性主轴）"))

    successful = [candidate for candidate in candidates if candidate.status == "成功"]
    if not successful:
        details = "；".join(candidate.status for candidate in candidates)
        return DimensionResult(source="近似最小包络", status=f"失败：{details}")
    return min(successful, key=lambda candidate: float(candidate.x) * float(candidate.y) * float(candidate.z))


def populate_approximate_envelopes(
    items: OrderedDict[str, BomItem],
    product_doc: Optional[Any],
    log: RunLog,
) -> None:
    """测量唯一零件的近似最小包络；每件后恢复总成，不切换 Part Design。"""
    for item in items.values():
        log.info("测量近似最小包络", item.part_number)
        try:
            item.envelope = measure_approximate_min_envelope(item.representative.part_doc)
            if item.envelope.status == "成功":
                log.info(
                    f"包络：{item.envelope.length:.3f} × {item.envelope.width:.3f} × {item.envelope.height:.3f} mm；{item.envelope.source}",
                    item.part_number,
                )
            else:
                log.warning(item.envelope.status, item.part_number)
        finally:
            if product_doc is not None:
                try:
                    product_doc.Activate()
                except Exception as exc:
                    log.warning(f"未能恢复总成文档：{exc}", item.part_number)


def convert_capture_to_original_png(capture_path: Path, image_path: Path) -> None:
    """无缩放转换 CATIA BMP 截图为 PNG；保留原始像素尺寸供 Excel 嵌入。"""
    with Image.open(capture_path) as image:
        if image.mode not in ("RGB", "L"):
            image = image.convert("RGB")
        image.save(image_path, "PNG", optimize=True)


def hide_stale_measurement_geometry(part_doc: Any, log: Optional[RunLog] = None, item_name: str = "") -> None:
    """删除本工具残留的临时测量几何，防止它影响截图的 Reframe。"""
    try:
        _remove_stale_temp_bodies(part_doc)
        if log:
            log.info("已清理残留临时测量对象", item_name)
    except Exception as exc:
        if log:
            log.warning(f"清理残留临时测量对象失败：{exc}", item_name)


def _select_all_leaf_products(root_product: Any, selection: Any) -> None:
    """递归加入总成树下所有叶子件，用于批量隐藏/恢复显示。"""
    products = root_product.Products
    for index in range(1, products.Count + 1):
        child = products.Item(index)
        if is_leaf(child):
            selection.Add(child)
        else:
            _select_all_leaf_products(child, selection)


def _select_leaf_products_under(target_product: Any, selection: Any) -> None:
    """选中目标节点下的所有叶子件；叶子件本身则只选中自身。"""
    if is_leaf(target_product):
        selection.Add(target_product)
    else:
        _select_all_leaf_products(target_product, selection)


def _set_all_leaf_visibility(product_doc: Any, show: bool) -> None:
    """在总成文档中批量显示或隐藏叶子件。"""
    selection = product_doc.Selection
    selection.Clear()
    _select_all_leaf_products(product_doc.Product, selection)
    if selection.Count:
        # catVisPropertyShowAttr=0；catVisPropertyNoShowAttr=1。
        selection.VisProperties.SetShow(0 if show else 1)
    selection.Clear()


def _com_vector(x: float, y: float, z: float) -> Any:
    """创建 CATIA Viewpoint3D 接口可接受的三维向量。"""
    return (float(x), float(y), float(z))


def set_standard_isometric_view(viewer: Any) -> None:
    """固定为统一等轴测视角，避免继承用户上一次操作的随机视角。"""
    viewpoint = viewer.Viewpoint3D
    # 固定斜向观察 + Z 轴朝上；所有零件使用同一视角定义。
    viewpoint.PutSightDirection(_com_vector(-1.0, -1.0, -1.0))
    viewpoint.PutUpDirection(_com_vector(0.0, 0.0, 1.0))
    viewer.Update()


def capture_node_image(
    catia: Any,
    target_product: Any,
    image_stem: str,
    original_doc: Any,
    log: RunLog,
    label: str,
    part_doc: Optional[Any] = None,
) -> tuple[Optional[Path], str]:
    """截图任意树节点：叶子件、子总成或根总成均可作为目标。"""
    capture_path = IMAGE_DIR / f"{image_stem}.bmp"
    image_path = IMAGE_DIR / f"{image_stem}.png"
    try:
        # 先隐藏此前运行尺寸脚本可能留下的辅助线，避免 Reframe 把零件缩成小点。
        if part_doc is not None:
            part_doc.Activate()
            hide_stale_measurement_geometry(part_doc, log, label)

        original_doc.Activate()
        _set_all_leaf_visibility(original_doc, show=False)

        selection = original_doc.Selection
        selection.Clear()
        # 对子总成显示其全部后代零件；对叶子件只显示自身。
        _select_leaf_products_under(target_product, selection)
        if selection.Count == 0:
            raise BomExportError("目标节点下没有可显示的叶子件")
        selection.VisProperties.SetShow(0)
        selection.Clear()

        viewer = catia.ActiveWindow.ActiveViewer
        # 统一等轴测，再以当前单件的包络居中并铺满画面。
        try:
            set_standard_isometric_view(viewer)
        except Exception as exc:
            log.warning(f"设置统一等轴测视角失败，将使用当前视角：{exc}", label)
        viewer.Update()
        time.sleep(0.2)
        viewer.Reframe()
        viewer.Update()
        time.sleep(0.4)
        # CaptureToFile 只截取 CATIA Viewer 场景，不受其他 Windows 窗口遮挡。
        viewer.CaptureToFile(BMP_CAPTURE_FORMAT, str(capture_path))
        if not capture_path.exists() or capture_path.stat().st_size == 0:
            raise BomExportError("CATIA 未生成截图文件")
        convert_capture_to_original_png(capture_path, image_path)
        # BMP 只作为无损截屏的中间文件；最终 PNG 保留全部原始像素并用于嵌入。
        capture_path.unlink(missing_ok=True)
        return image_path, "成功"
    except Exception as exc:
        status = f"失败：{exc}"
        log.error(status, label)
        return None, status
    finally:
        try:
            original_doc.Activate()
            # 导出完成后让总成恢复为全部叶子件显示，便于继续工作。
            _set_all_leaf_visibility(original_doc, show=True)
            catia.ActiveWindow.ActiveViewer.Update()
        except Exception as exc:
            log.warning(f"未能恢复总成显示状态：{exc}", label)


def capture_part_image(catia: Any, item: BomItem, original_doc: Any, log: RunLog) -> None:
    """截图唯一零件；相同零件号只调用一次。"""
    image_stem = clean_filename(f"零件_{item.part_number}_{item.revision or 'NA'}")
    item.image_path, item.image_status = capture_node_image(
        catia,
        item.representative.product,
        image_stem,
        original_doc,
        log,
        item.part_number,
        item.representative.part_doc,
    )


def capture_assembly_image(catia: Any, node: TreeNode, original_doc: Any, log: RunLog) -> None:
    """截图根总成或子总成，画面只保留该节点下的零件。"""
    image_stem = clean_filename(f"总成_{node.tree_number}_{node.part_number}")
    node.image_path, node.image_status = capture_node_image(
        catia,
        node.product,
        image_stem,
        original_doc,
        log,
        f"总成 {node.tree_number} {node.part_number}",
    )


def prepare_images_and_dimensions(catia: Any, original_doc: Any, items: OrderedDict[str, BomItem], log: RunLog) -> None:
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    total = len(items)
    for index, item in enumerate(items.values(), 1):
        label = f"{index}/{total}"
        log.info("测量 Create Rough Stock 尺寸", f"{label} {item.part_number}")
        if item.representative.part_doc is None:
            item.dimensions = DimensionResult(status="失败：无法取得 CATPart 文档")
            log.warning(item.dimensions.status, item.part_number)
        else:
            item.dimensions = measure_rough_stock_dimensions(item.representative.part_doc)
            if item.dimensions.status != "成功":
                log.error(item.dimensions.status, item.part_number)
        log.info("生成唯一件截图", f"{label} {item.part_number}")
        capture_part_image(catia, item, original_doc, log)


def style_cell(cell: Any, font: Font, fill: PatternFill, alignment: Alignment, border: Border) -> None:
    cell.font = font
    cell.fill = fill
    cell.alignment = alignment
    cell.border = border


def add_image(ws: Any, cell: str, image_path: Optional[Path]) -> Optional[tuple[int, int]]:
    if image_path is None or not image_path.exists():
        return None
    image = ExcelImage(str(image_path))
    scale = min(IMAGE_SIZE[0] / image.width, IMAGE_SIZE[1] / image.height, 1.0)
    image.width = round(image.width * scale)
    image.height = round(image.height * scale)
    ws.add_image(image, cell)
    return image.width, image.height


def configure_sheet(ws: Any, headers: list[str], title: str, widths: list[float]) -> tuple[Border, Font, PatternFill]:
    thin = Side(style="thin", color="B7C9D6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    title_font = Font(name="微软雅黑", bold=True, size=13, color="FFFFFF")
    title_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="微软雅黑", bold=True, size=10, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    last_col = get_column_letter(len(headers))
    ws.merge_cells(f"A1:{last_col}1")
    title_cell = ws["A1"]
    title_cell.value = title
    style_cell(title_cell, title_font, title_fill, center, border)
    for column in range(2, len(headers) + 1):
        style_cell(ws.cell(1, column), title_font, title_fill, center, border)
    ws.row_dimensions[1].height = 28

    for column, header in enumerate(headers, 1):
        cell = ws.cell(2, column, header)
        style_cell(cell, header_font, header_fill, center, border)
    ws.row_dimensions[2].height = 30
    for column, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(column)].width = width
    ws.freeze_panes = "A3"
    return border, Font(name="微软雅黑", size=10), PatternFill("solid", fgColor="FFFFFF")


def write_workbook(nodes: list[TreeNode], items: OrderedDict[str, BomItem], output: Path, root_name: str, log: RunLog) -> None:
    workbook = Workbook()
    tree_ws = workbook.active
    tree_ws.title = "产品结构"
    bom_ws = workbook.create_sheet("BOM汇总")
    log_ws = workbook.create_sheet("运行日志")

    tree_headers = [
        "层级", "树编号", "类型", "零件号", "版本", "零件名称", "实例名称", "父级路径", "数量", "图片", "X尺寸(mm)", "Y尺寸(mm)", "Z尺寸(mm)", "状态",
    ]
    border, data_font, white_fill = configure_sheet(
        tree_ws,
        tree_headers,
        f"产品结构：{root_name}",
        [8, 12, 10, 34, 10, 28, 32, 52, 8, 22, 14, 14, 14, 34],
    )
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    assembly_fill = PatternFill("solid", fgColor="D9EAF7")
    root_fill = PatternFill("solid", fgColor="9DC3E6")

    item_by_key = dict(items)
    for row, node in enumerate(nodes, 3):
        item = item_by_key.get(node.unique_key) if node.is_leaf else None
        type_label = "根产品" if node.level == 0 else ("零件" if node.is_leaf else "子总成")
        dimensions = item.dimensions if item else DimensionResult()
        status = f"尺寸：{dimensions.status}" if item else ""
        if item:
            status += f"；图片：{item.image_status}"
        values = [
            node.level, node.tree_number, type_label, node.part_number, node.revision,
            ("　" * node.level) + node.nomenclature, node.instance_name, node.parent_path,
            item.quantity if item else "", "", dimensions.x, dimensions.y, dimensions.z, status,
        ]
        fill = root_fill if node.level == 0 else (white_fill if node.is_leaf else assembly_fill)
        for column, value in enumerate(values, 1):
            alignment = left if column in (4, 6, 7, 8, 14) else center
            style_cell(tree_ws.cell(row, column, value), data_font, fill, alignment, border)
        if item and item.image_path:
            tree_ws.row_dimensions[row].height = IMAGE_SIZE[1] * 0.75 + 6
            add_image(tree_ws, f"J{row}", item.image_path)
        else:
            tree_ws.row_dimensions[row].height = 24
    if nodes:
        tree_ws.auto_filter.ref = f"A2:N{len(nodes) + 2}"

    bom_headers = [
        "序号", "零件号", "版本", "零件名称", "数量", "来源文件", "图片",
        "X尺寸(mm)", "Y尺寸(mm)", "Z尺寸(mm)", "长(mm)", "宽(mm)", "高(mm)", "尺寸口径", "尺寸状态", "图片状态",
    ]
    border, data_font, white_fill = configure_sheet(
        bom_ws,
        bom_headers,
        f"BOM 汇总（唯一件 {len(items)} 种）",
        [8, 34, 10, 28, 8, 42, 22, 14, 14, 14, 14, 14, 14, 32, 32, 32],
    )
    alternate_fill = PatternFill("solid", fgColor="EAF2F8")
    for seq, item in enumerate(items.values(), 1):
        dimension = item.dimensions
        row = seq + 2
        values = [
            seq, item.part_number, item.revision, item.nomenclature, item.quantity, item.source_file, "",
            dimension.x, dimension.y, dimension.z, dimension.length, dimension.width, dimension.height,
            dimension.source, dimension.status, item.image_status,
        ]
        fill = white_fill if seq % 2 else alternate_fill
        for column, value in enumerate(values, 1):
            alignment = left if column in (2, 4, 6, 14, 15, 16) else center
            style_cell(bom_ws.cell(row, column, value), data_font, fill, alignment, border)
        if item.image_path:
            bom_ws.row_dimensions[row].height = IMAGE_SIZE[1] * 0.75 + 6
            add_image(bom_ws, f"G{row}", item.image_path)
        else:
            bom_ws.row_dimensions[row].height = 24
    if items:
        bom_ws.auto_filter.ref = f"A2:P{len(items) + 2}"

    log_headers = ["级别", "对象", "信息"]
    border, data_font, white_fill = configure_sheet(log_ws, log_headers, "运行日志", [12, 42, 100])
    for index, (level, item_name, message) in enumerate(log.entries, 3):
        fill = PatternFill("solid", fgColor="FCE4D6") if level == "ERROR" else white_fill
        for column, value in enumerate((level, item_name, message), 1):
            style_cell(log_ws.cell(index, column, value), data_font, fill, left, border)
        log_ws.row_dimensions[index].height = 22
    if log.entries:
        log_ws.auto_filter.ref = f"A2:C{len(log.entries) + 2}"

    workbook.save(output)


def main() -> None:
    pythoncom.CoInitialize()
    log = RunLog()
    started = time.time()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        BOM_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        catia = connect_catia(log)
        product_doc, root_product = get_active_product_document(catia)
        root_name = com_text(root_product.PartNumber) or com_text(product_doc.Name)
        log.info("遍历产品树", root_name)
        nodes = traverse_product_tree(catia, root_product, log)
        items = create_bom_items(nodes)
        log.info(f"遍历完成：{len(nodes)} 个树节点，{len(items)} 种唯一零件")
        if not items:
            raise BomExportError("产品树没有可导出的叶子零件。")

        prepare_images_and_dimensions(catia, product_doc, items, log)
        output = OUTPUT_DIR / f"CATIA_BOM_{clean_filename(root_name)}_{timestamp}.xlsx"
        write_workbook(nodes, items, output, root_name, log)
        log.info(f"Excel 已生成：{output}")
        print(f"\n完成，用时 {time.time() - started:.1f} 秒。")
        print(f"Excel：{output}")
    except BomExportError as exc:
        log.error(str(exc))
        print(f"\n导出失败：{exc}")
    except Exception as exc:
        log.error(f"未处理异常：{exc}")
        traceback.print_exc()
        print(f"\n导出失败：{exc}")
    finally:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        log_path = LOG_DIR / f"CATIA_BOM_日志_{timestamp}.txt"
        log.save(log_path)
        print(f"日志：{log_path}")
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
