from __future__ import annotations

import json
import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from datatrans import config
from datatrans.job import run_job
from datatrans.headers import canonical_output_header, normalize_header

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data"
SAMPLE = DATA / "inputExample-total.xlsx"
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
                "export_eplan": options.get("export_eplan", True),
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
    def test_sample_generates_expected_tc_and_eplan(self) -> None:
        result = self.run_job_for([self.copy_sample()])
        self.assertEqual(result.status, "success")
        self.assertTrue(result.ok)
        self.assertTrue(Path(result.outputs["tc"]).exists())
        self.assertTrue(Path(result.outputs["eplan"]).exists())
        self.assertEqual(result.stats["rows_read"], 3)
        self.assertEqual(result.stats["tc_rows"], 3)
        self.assertEqual(result.stats["eplan_rows"], 3)

    def test_tc_values_and_blank_erp(self) -> None:
        result = self.run_job_for([self.copy_sample()])
        rows = self.load_rows(result.outputs["tc"])
        self.assertEqual(
            rows[0],
            ["是否合法", "部件编码", "名称1", "EPR编码", "类型编码", "制造商", "重量", "材料", "采购类型"],
        )
        self.assertEqual(
            rows[1],
            [None, "SHSS SYSX-X3585-2", "角度传感器", None, "SYS-X3585-2", "上海圣色", 0.1, "成品", "电器外购件"],
        )
        self.assertIsNone(rows[1][0])  # 是否合法不填
        self.assertIsNone(rows[1][3])  # EPR/ERP 编码不填
        self.assertEqual(rows[3][6], 115)
        self.assertEqual(rows[1][8], "电器外购件")

    def test_eplan_values_and_dimension_defaults(self) -> None:
        result = self.run_job_for([self.copy_sample()])
        rows = self.load_rows(result.outputs["eplan"])
        self.assertEqual(len(rows), 4)
        self.assertEqual(len(rows[0]), 28)
        self.assertEqual(rows[0][0], "部件编码")
        self.assertEqual(rows[1][0], "SHSS SYSX-X3585-2")
        self.assertEqual(rows[1][1], "角度传感器")
        self.assertEqual(rows[1][2], "SYS-X3585-2")
        self.assertEqual(rows[1][3], "件")
        self.assertIsNone(rows[1][4])  # 价格不填
        self.assertEqual(rows[1][5], "上海圣色")
        self.assertEqual(rows[1][6], "上海圣色")
        self.assertIsNone(rows[1][7])  # 描述不填
        self.assertEqual(rows[1][8], "1:电气工程")
        self.assertEqual(rows[1][11], "1:电气工程")
        self.assertEqual(rows[1][18:21], [100, 53, 300])
        # 第三行数据没有尺寸，默认 0
        self.assertEqual(rows[3][18:21], [0, 0, 0])
        for row in rows[1:]:
            self.assertEqual(row[8], "1:电气工程")
            self.assertEqual(row[11], "1:电气工程")

    def test_output_templates_keep_headers_and_remove_notes(self) -> None:
        result = self.run_job_for([self.copy_sample()])
        tc = load_workbook(result.outputs["tc"], data_only=True).active
        self.assertEqual(tc.max_row, 4)
        self.assertIsNone(tc["A11"].value)
        self.assertEqual(tc["D1"].value, "EPR编码")
        tc.parent.close()

        eplan = load_workbook(result.outputs["eplan"], data_only=True).active
        self.assertEqual(eplan.max_row, 4)
        self.assertEqual(eplan.max_column, 28)
        eplan.parent.close()

    def test_log_and_result_json_exist(self) -> None:
        result = self.run_job_for([self.copy_sample()])
        self.assertTrue(Path(result.log_path).exists())
        self.assertTrue(Path(result.result_path).exists())
        payload = json.loads(Path(result.result_path).read_text(encoding="utf-8"))
        self.assertEqual(payload["status"], "success")
        log_text = Path(result.log_path).read_text(encoding="utf-8-sig")
        self.assertIn("任务开始", log_text)
        self.assertIn("TC 表写出完成", log_text)


class TestValidation(BackendTestCase):
    def test_missing_required_fields_abort(self) -> None:
        path = self.copy_sample()
        workbook = load_workbook(path)
        worksheet = workbook.active
        worksheet["L4"] = None  # 备注
        worksheet["H5"] = None  # 代号
        workbook.save(path)
        workbook.close()

        result = self.run_job_for([path])
        self.assertEqual(result.status, "failed")
        self.assertFalse(result.ok)
        self.assertIsNone(result.outputs["tc"])
        self.assertIsNone(result.outputs["eplan"])
        codes = [issue.code for issue in result.issues]
        self.assertEqual(codes.count("REQUIRED_FIELD_EMPTY"), 2)
        self.assertFalse(any(self.output_dir.glob("TC_*.xlsx")))
        self.assertFalse(any(self.output_dir.glob("EPlan_*.xlsx")))

    def test_missing_required_fields_skip(self) -> None:
        path = self.copy_sample()
        workbook = load_workbook(path)
        worksheet = workbook.active
        worksheet["L4"] = None
        worksheet["H5"] = None
        workbook.save(path)
        workbook.close()

        result = self.run_job_for([path], on_invalid="skip")
        self.assertEqual(result.status, "partial")
        self.assertTrue(result.ok)
        self.assertEqual(result.stats["rows_invalid"], 2)
        rows = self.load_rows(result.outputs["tc"])
        self.assertEqual(len(rows), 2)  # 表头 + 1 条合法行
        self.assertEqual(rows[1][2], "举升电机")

    def test_unchecked_row_is_skipped_with_warning(self) -> None:
        path = self.copy_sample()
        workbook = load_workbook(path)
        worksheet = workbook.active
        worksheet["C4"] = None
        worksheet["D4"] = None
        workbook.save(path)
        workbook.close()

        result = self.run_job_for([path])
        self.assertEqual(result.status, "success")
        self.assertEqual(result.stats["rows_unchecked"], 1)
        self.assertIn("NO_EXPORT_TARGET", [issue.code for issue in result.issues])
        rows = self.load_rows(result.outputs["tc"])
        self.assertEqual(len(rows), 3)  # 表头 + 2 条勾选行

    def test_invalid_dimension_aborts_only_when_eplan_requested(self) -> None:
        path = self.copy_sample()
        workbook = load_workbook(path)
        worksheet = workbook.active
        worksheet["N4"] = "abc"
        workbook.save(path)
        workbook.close()

        result = self.run_job_for([path], export_tc=True, export_eplan=False)
        self.assertEqual(result.status, "success")
        self.assertNotIn("INVALID_NUMBER", [issue.code for issue in result.issues])

        result = self.run_job_for([path], export_tc=True, export_eplan=True)
        self.assertEqual(result.status, "failed")
        self.assertIn("INVALID_NUMBER", [issue.code for issue in result.issues])

    def test_only_xlsx_accepted(self) -> None:
        path = self.copy_sample("sample.xlsm")
        result = self.run_job_for([path])
        self.assertEqual(result.status, "failed")
        self.assertIn("UNSUPPORTED_EXTENSION", [issue.code for issue in result.issues])

    def test_missing_header_column(self) -> None:
        path = self.input_dir / "no_remark.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["名称", "代号", "部件编号", "EPLAN", "新TC"])
        worksheet.append(["测试", "C001", "P001", "√", "√"])
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
        self.assertEqual(result.stats["eplan_rows"], 6)
        self.assertEqual(
            sum(1 for issue in result.issues if issue.code == "DUPLICATE_ROW"), 3
        )
        rows = self.load_rows(result.outputs["tc"])
        self.assertEqual(len(rows), 7)


class TestHeaderHelpers(unittest.TestCase):
    def test_normalize_aliases(self) -> None:
        self.assertEqual(normalize_header("部件编号（厂家代号）"), "部件编号")
        self.assertEqual(normalize_header("重量（kg）"), "重量")
        self.assertEqual(normalize_header("额定电压（或工作电压范围）"), "额定电压")
        self.assertEqual(normalize_header("EPLan（√）"), "EPLAN")
        self.assertEqual(normalize_header("备注 (厂家)"), "备注")

    def test_output_alias_keeps_epr(self) -> None:
        self.assertEqual(canonical_output_header("EPR编码"), "ERP编码")
        self.assertEqual(canonical_output_header("ERP编码"), "ERP编码")
        self.assertEqual(canonical_output_header("行业/子行业"), "行业/子行业")
        self.assertEqual(canonical_output_header("宽度"), "宽度")


if __name__ == "__main__":
    unittest.main()
