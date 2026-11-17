from __future__ import annotations

import json
import shutil
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

from openpyxl import Workbook, load_workbook

from datatrans import config
from datatrans.headers import canonical_output_header, normalize_header
from datatrans.job import run_job

ROOT = Path(__file__).resolve().parents[2]
SAMPLE = ROOT / "backend" / "tests" / "fixtures" / "input_reference.xlsx"
TEMPLATES = ROOT / "templates"


class BackendTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory(prefix="datatrans_test_")
        self.base = Path(self.tmp.name)
        self.input_dir = self.base / "input"
        self.output_dir = self.base / "output"
        self.input_dir.mkdir()
        self.output_dir.mkdir()

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def copy_sample(self, name: str = "sample.xlsx") -> Path:
        target = self.input_dir / name
        shutil.copy2(SAMPLE, target)
        return target

    def run_job_for(self, inputs: list[Path], **options):
        payload = {
            "job_id": "testjob",
            "inputs": [str(path) for path in inputs],
            "output_dir": str(self.output_dir),
            "templates_dir": str(TEMPLATES),
            "options": {
                "export_tc": options.get("export_tc", True),
                "export_xml": options.get("export_xml", False),
                "on_invalid": options.get("on_invalid", "abort"),
            },
        }
        return run_job(payload)

    def load_rows(self, path: str | Path) -> list[list]:
        workbook = load_workbook(path, data_only=True)
        worksheet = workbook.active
        rows = [[cell.value for cell in row] for row in worksheet.iter_rows()]
        workbook.close()
        return rows


class TestSampleOutput(BackendTestCase):
    def test_hint_row_is_skipped(self) -> None:
        result = self.run_job_for([self.copy_sample()])
        self.assertEqual(result.status, "success")
        self.assertEqual(result.stats["rows_read"], 3)
        self.assertIn("HINT_ROW_SKIPPED", [issue.code for issue in result.issues])
        self.assertEqual([item.rows_valid for item in result.files], [3])

    def test_tc_output_mapping(self) -> None:
        result = self.run_job_for([self.copy_sample()])
        self.assertTrue(Path(result.outputs["tc"]).exists())
        rows = self.load_rows(result.outputs["tc"])
        self.assertEqual(
            rows[0],
            ["是否合法", "部件编码", "名称1", "EPR编码", "类型编码", "制造商", "重量", "材料", "采购类型"],
        )
        self.assertEqual(
            rows[1],
            [None, "国产.BLST-4.6-150", "扎带", "11111111", "BLST-4.6-150", "国产", 0.01, None, None],
        )
        self.assertEqual(len(rows), 4)
        self.assertIsNone(rows[1][0])
        self.assertIsNone(rows[1][7])
        self.assertIsNone(rows[1][8])

    def test_xml_output_mapping(self) -> None:
        result = self.run_job_for(
            [self.copy_sample()], export_tc=False, export_xml=True
        )
        self.assertEqual(result.status, "success")
        self.assertIsNone(result.outputs["tc"])
        self.assertEqual(result.stats["xml_rows"], 3)
        root = ET.parse(result.outputs["xml"]).getroot()
        self.assertEqual(root.tag, "partsmanagement")
        self.assertEqual(root.get("count"), "3")
        parts = root.findall("part")
        self.assertEqual(len(parts), 3)
        first = parts[0].attrib
        self.assertEqual(first["P_ARTICLE_PARTTYPE"], "1")
        self.assertEqual(first["P_ARTICLE_PARTNR"], "国产.BLST-4.6-150")
        self.assertEqual(first["P_ARTICLE_TYPENR"], "BLST-4.6-150")
        self.assertEqual(first["P_ARTICLE_ERPNR"], "11111111")
        self.assertEqual(first["P_ARTICLE_DESCR1"], "??_??@扎带;")
        self.assertTrue(first["P_ARTICLE_NOTE"].startswith("??_??@包塑型"))
        self.assertEqual(first["P_ARTICLE_WEIGHT"], "0.01")
        self.assertEqual(first["P_ARTICLE_WIDTH"], "4.6")
        self.assertEqual(first["P_ARTICLE_HEIGHT"], "250")
        self.assertEqual(first["P_ARTICLE_DEPTH"], "10")
        self.assertEqual(first["P_ARTICLE_MOUNTINGSPACE"], "1150")
        self.assertEqual(first["P_ARTICLE_MANUFACTURER"], "国产")
        self.assertEqual(first["P_ARTICLE_SUPPLIER"], "国产")
        self.assertEqual(first["P_ARTICLE_QUANTITYUNIT"], "??_??@条;")
        self.assertEqual(first["P_ARTICLE_PRODUCTGROUP"], "0")
        self.assertEqual(first["P_ARTICLE_PRODUCTSUBGROUP"], "0")

    def test_both_outputs(self) -> None:
        result = self.run_job_for([self.copy_sample()], export_tc=True, export_xml=True)
        self.assertEqual(result.status, "success")
        self.assertEqual(result.stats["tc_rows"], 3)
        self.assertEqual(result.stats["xml_rows"], 3)
        self.assertTrue(Path(result.outputs["tc"]).exists())
        self.assertTrue(Path(result.outputs["xml"]).exists())
        payload = json.loads(Path(result.result_path).read_text(encoding="utf-8"))
        self.assertEqual(payload["status"], "success")

    def test_log_and_result_json_exist(self) -> None:
        result = self.run_job_for([self.copy_sample()])
        self.assertTrue(Path(result.log_path).exists())
        self.assertTrue(Path(result.result_path).exists())
        log_text = Path(result.log_path).read_text(encoding="utf-8-sig")
        self.assertIn("任务开始", log_text)
        self.assertIn("TC 表写出完成", log_text)


class TestValidation(BackendTestCase):
    def test_required_fields(self) -> None:
        path = self.copy_sample()
        workbook = load_workbook(path)
        worksheet = workbook.active
        worksheet["B3"] = None
        worksheet["D3"] = None
        worksheet["E3"] = None
        worksheet["F3"] = None
        worksheet["G3"] = None
        workbook.save(path)
        workbook.close()

        result = self.run_job_for([path])
        self.assertEqual(result.status, "failed")
        self.assertEqual(result.stats["rows_invalid"], 1)
        self.assertEqual(
            sum(1 for issue in result.issues if issue.code == "REQUIRED_FIELD_EMPTY"), 5
        )
        self.assertIsNone(result.outputs["tc"])
        self.assertIsNone(result.outputs["xml"])

    def test_required_fields_skip_invalid(self) -> None:
        path = self.copy_sample()
        workbook = load_workbook(path)
        worksheet = workbook.active
        worksheet["B3"] = None
        workbook.save(path)
        workbook.close()

        result = self.run_job_for([path], on_invalid="skip")
        self.assertEqual(result.status, "partial")
        self.assertEqual(result.stats["rows_invalid"], 1)
        rows = self.load_rows(result.outputs["tc"])
        self.assertEqual(len(rows), 3)  # 表头 + 2 条合法行

    def test_negative_and_invalid_numbers(self) -> None:
        path = self.copy_sample()
        workbook = load_workbook(path)
        worksheet = workbook.active
        worksheet["J3"] = -1
        worksheet["K3"] = "abc"
        workbook.save(path)
        workbook.close()

        result = self.run_job_for([path])
        self.assertEqual(result.status, "failed")
        codes = [issue.code for issue in result.issues]
        self.assertIn("NEGATIVE_NUMBER", codes)
        self.assertIn("INVALID_NUMBER", codes)

    def test_blank_numbers_default_zero(self) -> None:
        path = self.copy_sample()
        workbook = load_workbook(path)
        worksheet = workbook.active
        for cell in ("J3", "K3", "L3", "M3"):
            worksheet[cell] = None
        workbook.save(path)
        workbook.close()

        result = self.run_job_for([path], export_tc=False, export_xml=True)
        self.assertEqual(result.status, "success")
        root = ET.parse(result.outputs["xml"]).getroot()
        first = root.find("part").attrib
        self.assertEqual(first["P_ARTICLE_WEIGHT"], "0")
        self.assertEqual(first["P_ARTICLE_WIDTH"], "0")
        self.assertEqual(first["P_ARTICLE_HEIGHT"], "0")
        self.assertEqual(first["P_ARTICLE_DEPTH"], "0")
        self.assertEqual(first["P_ARTICLE_MOUNTINGSPACE"], "0")

    def test_only_xlsx_accepted(self) -> None:
        path = self.copy_sample("sample.xlsm")
        result = self.run_job_for([path])
        self.assertEqual(result.status, "failed")
        self.assertIn("UNSUPPORTED_EXTENSION", [issue.code for issue in result.issues])

    def test_missing_header_column(self) -> None:
        path = self.input_dir / "no_supplier.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["部件编号", "类型编号", "名称1", "制造商"])
        worksheet.append(["P001", "C001", "测试", "厂家"])
        workbook.save(path)
        workbook.close()

        result = self.run_job_for([path])
        self.assertEqual(result.status, "failed")
        self.assertIn("INPUT_ERROR", [issue.code for issue in result.issues])


class TestMultiFile(BackendTestCase):
    def test_multiple_files_are_merged(self) -> None:
        first = self.copy_sample("a.xlsx")
        second = self.copy_sample("b.xlsx")
        result = self.run_job_for([first, second])
        self.assertEqual(result.status, "success")
        self.assertEqual(result.stats["files"], 2)
        self.assertEqual(result.stats["rows_read"], 6)
        self.assertEqual(result.stats["tc_rows"], 6)
        self.assertEqual(
            sum(1 for issue in result.issues if issue.code == "DUPLICATE_ROW"), 3
        )
        rows = self.load_rows(result.outputs["tc"])
        self.assertEqual(len(rows), 7)


class TestHeaderHelpers(unittest.TestCase):
    def test_normalize_aliases(self) -> None:
        self.assertEqual(normalize_header("类型编号"), "类型编号")
        self.assertEqual(normalize_header("名称1"), "名称1")
        self.assertEqual(normalize_header("重量（kg）"), "重量")
        self.assertEqual(normalize_header("宽度"), "宽度")

    def test_output_alias_keeps_epr(self) -> None:
        self.assertEqual(canonical_output_header("EPR编码"), "ERP编码")
        self.assertEqual(canonical_output_header("ERP编码"), "ERP编码")
        self.assertEqual(canonical_output_header("制造商"), "制造商")


if __name__ == "__main__":
    unittest.main()
