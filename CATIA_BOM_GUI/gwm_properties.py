# -*- coding: utf-8 -*-
"""长城汽车 CATIA 产品引用属性的读取与字段校验。"""

from __future__ import annotations

import re
from typing import Any


IDENTITY_FIELDS = frozenset(("V_description", "revision"))
PROPERTY_FIELDS = IDENTITY_FIELDS | frozenset((
    "E_XParamStrMatClassification", "E_XParamStrMatSpecifications",
    "E_XParamStr3DLWH", "gwmWeightUnit", "E_XParamStrActualWeight",
    "E_XParamStrDW", "E_XParamStrMatMark",
))


def _text(value: Any) -> str:
    try:
        if callable(value):
            value = value()
        return str(value).strip() if value is not None else ""
    except Exception:
        return ""


def _parameter_value(parameter: Any) -> str:
    for attribute in ("ValueAsString", "Value"):
        try:
            value = _text(getattr(parameter, attribute)).strip('"')
            if value:
                return value
        except Exception:
            continue
    return ""


def _parameters(product: Any, fields: frozenset[str]) -> dict[str, str]:
    """先读本节点自定义属性；不枚举包含全部后代的 Parameters。"""
    values: dict[str, str] = {}
    wanted = {name.casefold() for name in fields}
    try:
        collection = product.UserRefProperties
        count = int(collection.Count)
    except Exception:
        collection, count = None, 0
    for index in range(1, count + 1):
        try:
            parameter = collection.Item(index)
            name = _text(parameter.Name).split("\\")[-1].casefold()
            if name in wanted and not values.get(name):
                value = _parameter_value(parameter)
                if value:
                    values[name] = value
        except Exception:
            continue
        if wanted.issubset(values):
            return values

    try:
        parameters = product.Parameters
    except Exception:
        return values
    for name in sorted(fields):
        if name.casefold() in values:
            continue
        try:
            parameter = parameters.Item(name)
            full_name = _text(parameter.Name)
            if "\\" in full_name:
                # 按名查询也可能命中后代参数，只允许本节点自己的限定名。
                prefix, _tail = full_name.rsplit("\\", 1)
                owners = set()
                for attribute in ("Name", "PartNumber"):
                    try:
                        owners.add(_text(getattr(product, attribute)).casefold())
                    except Exception:
                        pass
                if prefix.casefold() not in owners:
                    continue
            if full_name.split("\\")[-1].casefold() != name.casefold():
                continue
            value = _parameter_value(parameter)
            if value:
                values[name.casefold()] = value
        except Exception:
            continue
    return values


def _number(value: str) -> float | None:
    match = re.fullmatch(r"\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*(?:mm)?\s*", value, re.I)
    return float(match.group(1)) if match else None


def read_properties(
    product: Any,
    reference_cache: dict | None = None,
    identity_only: bool = False,
) -> dict[str, Any]:
    """读取当前产品节点的引用属性；实例仅作为回退。"""
    try:
        reference = product.ReferenceProduct
    except Exception:
        reference = None
    fields = IDENTITY_FIELDS if identity_only else PROPERTY_FIELDS
    wanted = {name.casefold() for name in fields}
    values: dict[str, str] = {}
    seen = set()
    for candidate in (reference, product):
        if candidate is None:
            continue
        dispatch = getattr(candidate, "_oleobj_", None)
        identity = f"com:{dispatch!r}" if dispatch is not None else f"python:{id(candidate)}"
        if identity in seen:
            continue
        seen.add(identity)
        key = (identity, fields)
        if candidate is reference and reference_cache is not None:
            if key not in reference_cache:
                reference_cache[key] = _parameters(candidate, fields)
            attributes = reference_cache[key]
        else:
            attributes = _parameters(candidate, fields)
        for name, value in attributes.items():
            if value and not values.get(name):
                values[name] = value
        if wanted.issubset(values):
            break

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
