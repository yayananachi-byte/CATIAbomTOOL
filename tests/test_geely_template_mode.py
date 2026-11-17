"""Check Geely-only field mapping without changing the normal template mode."""

import importlib.util
import queue
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from openpyxl import Workbook


SCRIPT = Path(__file__).resolve().parents[1] / "模板转化" / "bom_to_control_table.py"
spec = importlib.util.spec_from_file_location("bom_to_control_table_geely_test", SCRIPT)
converter = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = converter
spec.loader.exec_module(converter)
sys.path.insert(0, str(SCRIPT.parents[1] / "CATIA_BOM_GUI"))
import catia_bom_gui as gui


class GeelyTemplateModeTests(unittest.TestCase):
    def test_standalone_conversion_uses_its_own_selection(self):
        for conversion_type, export_type in (("geely", "standard"), ("standard", "geely")):
            with self.subTest(conversion_type=conversion_type):
                app = SimpleNamespace(
                    template_bom_paths=["saved_bom.xlsx"],
                    template_file=Mock(get=Mock(return_value="template.xlsx")),
                    template_output=Mock(get=Mock(return_value="output.xlsx")),
                    conversion_type=Mock(get=Mock(return_value=conversion_type)),
                    data_type=Mock(get=Mock(return_value=export_type)),
                    _start_worker=Mock(),
                )
                gui.BomManagerApp._start_conversion(app)
                action, payload = app._start_worker.call_args.args
                self.assertEqual(action, "convert")
                self.assertEqual(payload["data_type"], conversion_type)
                app.data_type.get.assert_not_called()

    def test_names_go_to_the_matching_number_column_and_q_stays_blank(self):
        sheet = Workbook().active
        for row, kind, path, name, column in (
            (10, "根产品", ["1"], "根总成/001", "M"),
            (16, "子总成", ["1", "1"], "分总成/002", "N"),
            (22, "零件", ["1", "1", "1"], "叶子件/003", "O"),
        ):
            record = {"类型": kind, "零件号": f"P{row}", "零件名称": name,
                      "_tree_no": ".".join(path), "_path": path}
            converter.write_record(sheet, row, record, None, geely_mode=True)
            self.assertEqual(sheet[f"{column}{row}"].value, name.split("/")[0])
            self.assertIsNone(sheet[f"Q{row}"].value)

    def test_normal_mode_keeps_part_number_and_name(self):
        sheet = Workbook().active
        record = {"类型": "零件", "零件号": "P22", "零件名称": "叶子件",
                  "_tree_no": "1.1", "_path": ["1", "1"]}
        converter.write_record(sheet, 10, record, None)
        self.assertEqual(sheet["O10"].value, "P22")
        self.assertEqual(sheet["Q10"].value, "叶子件")

    def test_one_click_passes_geely_mode_to_conversion(self):
        source = SCRIPT.parents[1] / "模板转化" / "CATIA_BOM_LCT-80010514_20261019_195605.xlsx"
        template = SCRIPT.parents[1] / "模板转化" / "客户XXX-项目名XXX-基础数据管制表-日期.xlsx"
        worker = gui.GuiWorker("export_and_convert", {
            "export": {"data_type": "geely"},
            "template_path": str(template),
            "output_path": str(SCRIPT.parents[1] / "tests" / "geely_output.xlsx"),
        }, queue.Queue())
        with patch.object(worker, "_run_export", return_value={
            "results": [str(source)], "output_dir": str(source.parent),
        }), patch.object(worker, "_run_convert", return_value={
            "results": ["converted.xlsx"], "output_dir": str(source.parent),
        }) as run_convert:
            worker._run_export_and_convert()
        self.assertEqual(run_convert.call_args.args[0]["data_type"], "geely")

    def test_converter_receives_geely_mode(self):
        source = SCRIPT.parents[1] / "模板转化" / "CATIA_BOM_LCT-80010514_20261019_195605.xlsx"
        template = SCRIPT.parents[1] / "模板转化" / "客户XXX-项目名XXX-基础数据管制表-日期.xlsx"
        converter_mock = SimpleNamespace(convert_many=Mock(), ConversionCancelled=RuntimeError)
        worker = gui.GuiWorker("convert", {
            "bom_paths": [str(source)], "template_path": str(template),
            "output_path": str(SCRIPT.parents[1] / "tests" / "geely_output.xlsx"),
            "data_type": "geely",
        }, queue.Queue())
        with patch.object(gui, "select_converter", return_value=(converter_mock, "基础数据管制表模板")):
            worker._run_convert()
        self.assertTrue(converter_mock.convert_many.call_args.kwargs["geely_mode"])


if __name__ == "__main__":
    unittest.main()
