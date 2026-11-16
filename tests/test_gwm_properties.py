"""长城属性读取的离线回归测试；不连接 CATIA。"""

from collections import Counter
from pathlib import Path
import sys
import unittest


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "CATIA_BOM_GUI"))
from gwm_properties import read_properties


class Parameter:
    def __init__(self, name, value, *, text_error=False):
        self.Name = name
        self.value = value
        self.text_error = text_error
        self.calls = Counter()

    @property
    def ValueAsString(self):
        self.calls["ValueAsString"] += 1
        if self.text_error:
            raise RuntimeError("ValueAsString unavailable")
        return self.value

    @property
    def Value(self):
        self.calls["Value"] += 1
        return self.value


class UserProperties:
    def __init__(self, parameters=()):
        self.parameters = list(parameters)
        self.calls = Counter()

    @property
    def Count(self):
        self.calls["Count"] += 1
        return len(self.parameters)

    def Item(self, index):
        self.calls["Item"] += 1
        return self.parameters[index - 1]


class AggregateParameters:
    """模拟 Product.Parameters：可按名返回子节点参数，禁止全量枚举。"""

    def __init__(self, named=None):
        self.named = named or {}
        self.calls = Counter()

    @property
    def Count(self):
        self.calls["Count"] += 1
        raise AssertionError("Aggregated descendant parameters must not be enumerated")

    def Item(self, name):
        self.calls["Item"] += 1
        if not isinstance(name, str):
            raise AssertionError("Only named parameter queries are allowed")
        return self.named[name]


class Product:
    def __init__(self, values=None, *, reference=None, name="Product", part_number="PN"):
        self.Name = name
        self.PartNumber = part_number
        self.ReferenceProduct = reference if reference is not None else self
        self.UserRefProperties = UserProperties(
            Parameter(key, value) for key, value in (values or {}).items()
        )
        self.Parameters = AggregateParameters()


def sample_values():
    return {
        "V_description": "右扭力盒",
        "revision": "G.1",
        "E_XParamStrMatClassification": "钢板",
        "E_XParamStrMatSpecifications": "2.0",
        "E_XParamStrMatMark": "HC260YD+ZM 35/35-M",
        "E_XParamStr3DLWH": "397.6*313.1*204.9",
        "E_XParamStrDW": "1.6268",
        "E_XParamStrActualWeight": "0.0",
        "gwmWeightUnit": "kg",
    }


class GwmPropertyTests(unittest.TestCase):
    def test_screenshot_fields_and_zero_actual_weight_fall_back_to_design(self):
        result = read_properties(Product(sample_values()))
        self.assertEqual(result["name"], "右扭力盒")
        self.assertEqual(result["revision"], "G.1")
        self.assertEqual(result["material"], "HC260YD+ZM 35/35-M")
        self.assertEqual(result["thickness_mm"], 2.0)
        self.assertEqual(result["dimensions"], (397.6, 313.1, 204.9))
        self.assertEqual(result["weight_kg"], 1.6268)
        self.assertIn("E_XParamStrDW", result["weight_note"])

    def test_empty_thickness_and_actual_weight(self):
        values = sample_values()
        values.update(E_XParamStrMatSpecifications="", E_XParamStrActualWeight="")
        result = read_properties(Product(values))
        self.assertIsNone(result["thickness_mm"])
        self.assertEqual(result["weight_kg"], 1.6268)

    def test_actual_weight_precedes_design_and_units_are_converted(self):
        for unit, raw, expected in (("kg", "1.5", 1.5), ("g", "1500", 1.5)):
            with self.subTest(unit=unit):
                values = sample_values()
                values.update(gwmWeightUnit=unit, E_XParamStrActualWeight=raw)
                result = read_properties(Product(values))
                self.assertEqual(result["weight_kg"], expected)
                self.assertEqual(result["weight_note"], "E_XParamStrActualWeight")

    def test_missing_weight_unit_does_not_guess_kg(self):
        values = sample_values()
        values["gwmWeightUnit"] = ""
        self.assertIsNone(read_properties(Product(values))["weight_kg"])

    def test_reference_priority_with_instance_fallback(self):
        reference = Product({"V_description": "引用名称", "revision": ""})
        instance = Product(
            {"V_description": "实例名称", "revision": "B.1"}, reference=reference
        )
        result = read_properties(instance)
        self.assertEqual(result["name"], "引用名称")
        self.assertEqual(result["revision"], "B.1")

    def test_value_getter_used_if_value_as_string_getter_raises(self):
        product = Product()
        parameter = Parameter("V_description", "回退名称", text_error=True)
        product.UserRefProperties = UserProperties([parameter])
        self.assertEqual(read_properties(product)["name"], "回退名称")
        self.assertEqual(parameter.calls["Value"], 1)

    def test_product_custom_properties_beat_aggregated_descendant(self):
        product = Product({"V_description": "根总成"}, name="Root")
        child_name = Parameter("Child\\V_description", "子零件")
        product.Parameters = AggregateParameters({"V_description": child_name})
        self.assertEqual(read_properties(product)["name"], "根总成")
        self.assertEqual(child_name.calls["ValueAsString"], 0)
        self.assertEqual(product.Parameters.calls["Count"], 0)

    def test_absent_root_property_does_not_leak_descendant_name(self):
        product = Product(name="Root", part_number="RootPN")
        child_name = Parameter("Child\\V_description", "子零件")
        product.Parameters = AggregateParameters({"V_description": child_name})
        self.assertEqual(read_properties(product)["name"], "")
        self.assertEqual(child_name.calls["ValueAsString"], 0)
        self.assertEqual(product.Parameters.calls["Count"], 0)

    def test_scoped_parameter_fallback_accepts_current_product(self):
        product = Product(name="Root")
        product.Parameters = AggregateParameters(
            {"V_description": Parameter("Root\\V_description", "根总成")}
        )
        self.assertEqual(read_properties(product)["name"], "根总成")

    def test_qualified_custom_property_name_and_case_are_supported(self):
        product = Product({"Root\\v_DESCRIPTION": "根总成"})
        self.assertEqual(read_properties(product)["name"], "根总成")

    def test_identity_only_never_reads_material_values(self):
        product = Product(sample_values())
        result = read_properties(product, identity_only=True)
        self.assertEqual(result["name"], "右扭力盒")
        self.assertEqual(result["revision"], "G.1")
        material = next(p for p in product.UserRefProperties.parameters if "MatMark" in p.Name)
        self.assertEqual(material.calls["ValueAsString"], 0)
        self.assertEqual(material.calls["Value"], 0)

    def test_reference_cache_reused_even_when_reference_has_missing_fields(self):
        reference = Product({"V_description": "共同引用"})
        first = Product({"revision": "A.1"}, reference=reference)
        second = Product({"revision": "B.1"}, reference=reference)
        cache = {}
        self.assertEqual(read_properties(first, cache)["revision"], "A.1")
        reference_reads = reference.UserRefProperties.calls["Count"]
        self.assertEqual(read_properties(second, cache)["revision"], "B.1")
        self.assertEqual(reference.UserRefProperties.calls["Count"], reference_reads)
        self.assertEqual(reference_reads, 1)

    def test_new_export_cache_rereads_reference(self):
        reference = Product({"V_description": "第一次"})
        instance = Product(reference=reference)
        self.assertEqual(read_properties(instance, {})["name"], "第一次")
        reference.UserRefProperties.parameters[0].value = "第二次"
        self.assertEqual(read_properties(instance, {})["name"], "第二次")
        self.assertEqual(reference.UserRefProperties.calls["Count"], 2)

    def test_different_references_with_same_part_number_do_not_share_values(self):
        first = Product({"V_description": "版本甲"}, part_number="SAME")
        second = Product({"V_description": "版本乙"}, part_number="SAME")
        cache = {}
        self.assertEqual(read_properties(Product(reference=first), cache)["name"], "版本甲")
        self.assertEqual(read_properties(Product(reference=second), cache)["name"], "版本乙")

    def test_identity_and_full_reads_have_separate_cache_entries(self):
        product = Product(sample_values())
        cache = {}
        read_properties(product, cache, identity_only=True)
        self.assertEqual(read_properties(product, cache)["thickness_mm"], 2.0)

    def test_same_reference_and_instance_are_not_scanned_twice(self):
        product = Product({"V_description": "自引用"})
        read_properties(product)
        self.assertEqual(product.UserRefProperties.calls["Count"], 1)


if __name__ == "__main__":
    unittest.main()
