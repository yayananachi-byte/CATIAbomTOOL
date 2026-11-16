"""Exercise the shared export pipeline without changing CATIA documents."""

import sys
import tempfile
import unittest
from contextlib import ExitStack
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "CATIA_BOM_GUI"))
import catia_bom_export_v2 as core


class ExportModeTests(unittest.TestCase):
    def export(self, mode, duplicate_part=False):
        nodes = []
        for index, (name, leaf) in enumerate((("Root/A", False), ("Sub/B", False), ("Part/C", True))):
            product = SimpleNamespace(PartNumber=f"P{index}")
            node = core.TreeNode(index, f"1{'.1' * index}", "", "", f"Instance{index}",
                                 f"P{index}", "R1", name, leaf, product,
                                 part_doc=object() if leaf else None)
            nodes.append(node)
        if duplicate_part:
            nodes.append(core.TreeNode(2, "1.2.1", "", "", "OtherInstance", "P2", "R2",
                                      "Other part", True, SimpleNamespace(PartNumber="P2"),
                                      part_doc=object()))
        details = {
            "name": "GWM name", "revision": "G.1", "material": "Steel",
            "thickness_mm": 2.0, "thickness_note": "Property thickness",
            "weight_kg": 1.6268, "weight_note": "Design weight",
            "dimensions": (397.6, 313.1, 204.9),
        }
        document = SimpleNamespace(Product=nodes[0].product, Name="test.CATProduct", Activate=Mock())
        methods = ("populate_materials", "populate_nominal_thicknesses", "populate_masses",
                   "populate_rough_stock_dimensions", "hide_stale_measurement_geometry",
                   "capture_assembly_image", "capture_part_image", "write_workbook",
                   "_restore_product_after_export", "_cleanup_capture_temp_files")
        with ExitStack() as stack:
            temporary = stack.enter_context(tempfile.TemporaryDirectory())
            stack.enter_context(patch.object(core, "traverse_product_tree", return_value=nodes))
            mocks = {name: stack.enter_context(patch.object(core, name)) for name in methods}
            state = stack.enter_context(patch.object(core, "ScreenshotDisplayState")).return_value
            read = stack.enter_context(patch("gwm_properties.read_properties", return_value=details))
            if duplicate_part:
                read.side_effect = [details, details, details, dict(details, weight_kg=9.0)]
            result = core.export_product_document(
                object(), document, Path(temporary) / "bom.xlsx", Mock(),
                image_directory=Path(temporary) / "images", data_type=mode)
            self.assertEqual(result, "P0")
            written_nodes, items = mocks["write_workbook"].call_args.args[:2]
            self.assertEqual(mocks["capture_assembly_image"].call_count, 2)
            self.assertEqual(mocks["capture_part_image"].call_count, 1)
            state.prepare.assert_called_once()
            state.restore.assert_called_once()
            return written_nodes, items, mocks, read

    def test_standard_measurements_and_names_preserved(self):
        nodes, _, mocks, read = self.export("standard")
        self.assertEqual([n.nomenclature for n in nodes], ["Root/A", "Sub/B", "Part/C"])
        read.assert_not_called()
        for name in ("populate_materials", "populate_nominal_thicknesses", "populate_masses",
                     "populate_rough_stock_dimensions"):
            mocks[name].assert_called_once()

    def test_geely_trims_all_levels_and_preserves_measurements(self):
        nodes, items, mocks, read = self.export("geely")
        self.assertEqual([n.nomenclature for n in nodes], ["Root", "Sub", "Part"])
        self.assertEqual([n.revision for n in nodes], ["R1"] * 3)
        self.assertEqual(items["P2"].nomenclature, "Part")
        read.assert_not_called()
        for name in ("populate_materials", "populate_nominal_thicknesses", "populate_masses",
                     "populate_rough_stock_dimensions"):
            mocks[name].assert_called_once()

    def test_gwm_names_all_levels_and_only_reads_properties(self):
        nodes, items, mocks, read = self.export("gwm")
        self.assertEqual([n.nomenclature for n in nodes], ["GWM name"] * 3)
        self.assertEqual([call.kwargs["identity_only"] for call in read.call_args_list], [True, True, False])
        self.assertTrue(all(call.args[1] is read.call_args.args[1] for call in read.call_args_list))
        part = items["P2"]
        self.assertEqual(part.mass.value_kg, 1.6268)
        self.assertEqual(part.nominal_thickness.value, "2 mm")
        self.assertEqual(part.dimensions.length, 397.6)
        for name in ("populate_materials", "populate_nominal_thicknesses", "populate_masses",
                     "populate_rough_stock_dimensions"):
            mocks[name].assert_not_called()

    def test_same_number_gwm_item_keeps_its_first_representative_attributes(self):
        _, items, _, _ = self.export("gwm", duplicate_part=True)
        self.assertEqual(items["P2"].quantity, 2)
        self.assertEqual(items["P2"].mass.value_kg, 1.6268)

    def test_prepared_capture_session_avoids_repeated_part_activation_and_cleanup(self):
        for prepared in (True, False):
            with self.subTest(prepared=prepared), ExitStack() as stack:
                part_doc = SimpleNamespace(Activate=Mock())
                original = SimpleNamespace(Activate=Mock())
                viewer = SimpleNamespace(Update=Mock(), Reframe=Mock(), CaptureToFile=Mock())
                catia = SimpleNamespace(ActiveWindow=SimpleNamespace(ActiveViewer=viewer))
                state = Mock(prepared=prepared)
                state.show_only.return_value = [object()]
                cleanup = stack.enter_context(patch.object(core, "hide_stale_measurement_geometry"))
                stack.enter_context(patch.object(core, "set_standard_isometric_view"))
                stack.enter_context(patch.object(core, "convert_capture_to_original_png"))
                stack.enter_context(patch.object(core.time, "sleep"))
                stack.enter_context(patch.object(Path, "exists", return_value=True))
                stack.enter_context(patch.object(Path, "stat", return_value=SimpleNamespace(st_size=1)))
                stack.enter_context(patch.object(Path, "unlink"))
                _, status = core.capture_node_image(catia, object(), "test", original, Mock(),
                                                    "part", part_doc, state, [object()])
                self.assertEqual(status, "成功")
                self.assertEqual(part_doc.Activate.call_count, 0 if prepared else 1)
                self.assertEqual(cleanup.call_count, 0 if prepared else 1)
                viewer.Reframe.assert_called_once()
                viewer.CaptureToFile.assert_called_once()


if __name__ == "__main__":
    unittest.main()
