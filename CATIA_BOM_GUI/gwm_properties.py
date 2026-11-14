# -*- coding: utf-8 -*-
"""长城汽车 CATIA 产品引用属性的读取与字段校验。"""

from __future__ import annotations

import re
from typing import Any


def _text(value: Any) -> str:
    try:
        if callable(value):
            value = value()
        return str(value).strip() if value is not None else ""
    except Exception:
        return ""


def _parameters(product: Any) -> dict[str, str]:
    values: dict[str, str] = {}
    for collection_name in ("Parameters", "UserRefProperties"):
        try:
            collection = getattr(product, collection_name)
            count = int(collection.Count)
        except Exception:
            continue
        for index in range(1, count + 1):
            try:
                parameter = collection.Item(index)
                name = _text(parameter.Name).split("\\")[-1].casefold()
                if name and name not in values:
                    value = _text(getattr(parameter, "ValueAsString", None))
                    if not value:
                        value = _text(getattr(parameter, "Value", None))
                    values[name] = value.strip('"')
            except Exception:
                continue
    return values


def _number(value: str) -> float | None:
    match = re.fullmatch(r"\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*(?:mm)?\s*", value, re.I)
    return float(match.group(1)) if match else None


def read_properties(product: Any) -> dict[str, Any]:
    """读取当前产品节点的引用属性；实例仅作为回退。"""
    try:
        reference = product.ReferenceProduct
    except Exception:
        reference = None
    values: dict[str, str] = {}
    for candidate in (reference, product):
        if candidate is None:
            continue
        for name, value in _parameters(candidate).items():
            if value and not values.get(name):
                values[name] = value

    def field(name: str) -> str:
        return values.get(name.casefold(), "")

    classification = field("E_XParamStrMatClassification")
    specification = field("E_XParamStrMatSpecifications")
    thickness = _number(specification) if "钢板" in classification else None
    dimensions_text = field("E_XParamStr3DLWH")
    dimensions = None
    parts = re.split(r"[xX×*＊]", dimensions_text)
    if len(parts) == 3:
        parsed = [_number(part) for part in parts]
        if all(number is not None and number > 0 for number in parsed):
            dimensions = tuple(parsed)

    unit = field("gwmWeightUnit").casefold()
    actual = _number(field("E_XParamStrActualWeight"))
    design = _number(field("E_XParamStrDW"))
    weight = actual if actual is not None and actual > 0 else design
    if weight is not None and weight > 0 and unit in ("kg", "千克"):
        weight_kg = weight
    elif weight is not None and weight > 0 and unit in ("g", "克"):
        weight_kg = weight / 1000
    else:
        weight_kg = None

    return {
        "name": field("V_description"),
        "revision": field("revision"),
        "material": field("E_XParamStrMatMark"),
        "thickness_mm": thickness if thickness is not None and thickness > 0 else None,
        "thickness_note": (
            "E_XParamStrMatSpecifications（钢板，需图纸确认）"
            if thickness is not None and thickness > 0 else "属性未提供可确认的钢板厚度"
        ),
        "weight_kg": weight_kg,
        "weight_note": "E_XParamStrActualWeight" if actual is not None and actual > 0 else "E_XParamStrDW（设计重量）",
        "dimensions": dimensions,
    }
