"""读取 .xlsx 输入，定位表头并产出原始行。

支持两种数据来源：
- openpyxl 直接读文件（默认，快）；
- Excel COM 读单元格（用于内网透明加密文件，见 excel_com.py）。

两者最终都转成 “网格 + 表头” 的同一套逻辑，避免校验/映射出现两套代码。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from . import config
from .headers import HeaderInfo, describe_header, find_input_header, normalize_header
from .models import Issue, RawRow


class InputError(Exception):
    """输入文件格式/可读性错误。"""


@dataclass(slots=True)
class ReadResult:
    file: str
    sheet: str
    header: HeaderInfo | None
    rows: list[RawRow] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    engine: str = "openpyxl"


@dataclass(slots=True)
class PdmReadResult:
    """PDM 输入：不按 TC 表头解析，直接按 B:N 原样取行。"""

    file: str
    sheet: str
    rows: list[list[Any]] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    engine: str = "openpyxl"


class _Cell:
    __slots__ = ("value",)

    def __init__(self, value: Any = None) -> None:
        self.value = value


class MatrixSheet:
    """把 Excel COM 返回的二维数组包装成 openpyxl Worksheet 的最小接口。"""

    def __init__(
        self,
        values: list[list[Any]],
        *,
        first_row: int = 1,
        first_col: int = 1,
        title: str = "Sheet1",
    ) -> None:
        self._values = values or []
        self._first_row = max(1, int(first_row))
        self._first_col = max(1, int(first_col))
        self.title = title
        row_count = len(self._values)
        col_count = max((len(row) for row in self._values), default=0)
        self.max_row = self._first_row + row_count - 1 if row_count else self._first_row
        self.max_column = self._first_col + col_count - 1 if col_count else self._first_col

    def cell(self, row: int, column: int = 1) -> _Cell:
        r = int(row) - self._first_row
        c = int(column) - self._first_col
        if 0 <= r < len(self._values):
            row_values = self._values[r]
            if 0 <= c < len(row_values):
                return _Cell(row_values[c])
        return _Cell(None)


def _issue(
    level: str,
    code: str,
    message: str,
    file: str,
    sheet: str = "",
    row: int | None = None,
    field_name: str = "",
) -> Issue:
    return Issue(
        level=level,
        code=code,
        message=message,
        file=file,
        sheet=sheet,
        row=row,
        field=field_name,
    )


def _is_note_row(grid: Any, row_idx: int, values: dict[str, Any]) -> bool:
    """表格末尾的“注意”说明行：缺少类型编号/部件编号，但某个单元格含“注意”。

    要求核心字段至少缺一个，避免把备注里写了“注意安装方向”的正常物料行误删。
    """
    if not any(_is_blank(values.get(field)) for field in config.NOTE_ROW_CORE_FIELDS):
        return False
    for col_idx in range(1, grid.max_column + 1):
        value = grid.cell(row=row_idx, column=col_idx).value
        if value is None:
            continue
        text = str(value)
        if any(keyword in text for keyword in config.NOTE_ROW_KEYWORDS):
            return True
    return False


def read_grid(
    grid: Any,
    *,
    file_text: str,
    sheet_title: str,
    engine: str,
    extra_issues: list[Issue] | None = None,
) -> ReadResult:
    """从任意支持 max_row/max_column/cell() 的网格读取数据。"""
    result = ReadResult(
        file=file_text,
        sheet=sheet_title,
        header=None,
        issues=list(extra_issues or []),
        engine=engine,
    )

    header = find_input_header(grid)
    if header is None:
        raise InputError(
            "找不到表头行。请确认输入模板包含“部件编号、类型编号、名称1、制造商、供应商”等列。"
        )
    result.header = header

    if header.duplicates:
        result.issues.append(
            _issue(
                "warning",
                "DUPLICATE_HEADER",
                "存在重复表头，已使用第一次出现的列：" + "、".join(header.duplicates),
                file_text,
                sheet_title,
                header.row,
            )
        )

    missing = header.missing_required
    if missing:
        raise InputError(
            "缺少必需列：" + "、".join(missing) + "。识别到的表头：" + describe_header(header)
        )

    optional_defaults = {
        "ERP编码": "留空",
        "描述": "留空",
        "数量单位": config.DEFAULT_UNIT,
        "重量": "0",
        "宽": "0",
        "高": "0",
        "深": "0",
    }
    for name, default in optional_defaults.items():
        if name not in header.columns:
            result.issues.append(
                _issue(
                    "warning",
                    "MISSING_OPTIONAL_HEADER",
                    f"缺少可选列“{name}”，将使用默认值：{default}",
                    file_text,
                    sheet_title,
                    header.row,
                    name,
                )
            )

    # 逐行读取，不因中间空行提前停止；全空行跳过。
    for row_idx in range(header.data_start_row, grid.max_row + 1):
        values: dict[str, Any] = {}
        for canonical, col_idx in header.columns.items():
            values[canonical] = grid.cell(row=row_idx, column=col_idx).value
        if all(_is_blank(value) for value in values.values()):
            continue
        # 新模板第 2 行是提示内容（通常没有序号），跳过，不算数据。
        if (
            row_idx == header.data_start_row
            and "序号" in header.columns
            and _is_blank(values.get("序号"))
        ):
            result.issues.append(
                _issue(
                    "info",
                    "HINT_ROW_SKIPPED",
                    f"第 {row_idx} 行是模板提示内容，已跳过",
                    file_text,
                    sheet_title,
                    row_idx,
                )
            )
            continue
        if _is_note_row(grid, row_idx, values):
            result.issues.append(
                _issue(
                    "info",
                    "NOTE_ROW_SKIPPED",
                    f"第 {row_idx} 行包含“注意”说明且没有核心字段，已按说明行跳过",
                    file_text,
                    sheet_title,
                    row_idx,
                )
            )
            continue
        result.rows.append(
            RawRow(
                file=file_text,
                sheet=sheet_title,
                row=row_idx,
                values=values,
            )
        )

    if not result.rows:
        result.issues.append(
            _issue(
                "warning",
                "NO_DATA_ROWS",
                "表头下方没有读到数据行",
                file_text,
                sheet_title,
                header.data_start_row,
            )
        )

    return result


def read_input_file(
    path: str | Path,
    sheet_name: str | None = None,
) -> ReadResult:
    """用 openpyxl 直接读取 .xlsx（普通文件路径）。"""
    file_path = Path(path)
    file_text = str(file_path)

    if file_path.suffix.lower() not in config.INPUT_EXTENSIONS:
        raise InputError(f"只支持 {config.INPUT_EXTENSION_LABEL} 格式：{file_path.name}")
    if not file_path.exists():
        raise InputError(f"文件不存在：{file_path}")
    if not file_path.is_file():
        raise InputError(f"不是文件：{file_path}")

    try:
        workbook = load_workbook(
            filename=file_path,
            data_only=True,
            read_only=False,
            keep_links=False,
        )
    except PermissionError as exc:
        raise InputError(f"无法读取文件（可能被占用）：{file_path}") from exc
    except InvalidFileException as exc:
        raise InputError(f"不是有效的 Excel .xlsx 文件：{file_path}") from exc
    except Exception as exc:  # openpyxl 对损坏/加密文件会抛多种异常
        raise InputError(f"openpyxl 打开失败：{file_path}（{exc}）") from exc

    try:
        if sheet_name:
            if sheet_name not in workbook.sheetnames:
                raise InputError(f"工作表不存在：{sheet_name}")
            worksheet = workbook[sheet_name]
            extra_issues: list[Issue] = []
        else:
            worksheet = workbook.active or workbook.worksheets[0]
            extra_issues = []
            if len(workbook.sheetnames) > 1:
                extra_issues.append(
                    _issue(
                        "info",
                        "MULTI_SHEET",
                        f"工作簿包含多个工作表，使用活动工作表“{worksheet.title}”",
                        file_text,
                        worksheet.title,
                    )
                )

        return read_grid(
            worksheet,
            file_text=file_text,
            sheet_title=worksheet.title,
            engine="openpyxl",
            extra_issues=extra_issues,
        )
    finally:
        workbook.close()


_PDM_HEADER_LABELS = {
    normalize_header(value)
    for value in (
        "部件编号",
        "主数据编码",
        "零件组ID",
        "长描述",
        "查找编号",
        "bl_item_图号",
        "版本名称",
        "bl_rev_规格",
        "bl_rev_材料",
        "bl_rev_件重",
        "数量",
        "bl_rev_备注",
        "备注1",
    )
}


def pdm_rows_from_matrix(matrix: list[list[Any]]) -> list[list[Any]]:
    """从二维矩阵取 B:N（第 2~14 列），第一行若像 PDM 表头则跳过。"""
    rows: list[list[Any]] = []
    header_checked = False
    for raw in matrix:
        values = list(raw or [])
        payload = list(values[1:14])
        payload.extend([None] * (13 - len(payload)))
        if all(_is_blank(value) for value in payload):
            continue
        if not header_checked:
            header_checked = True
            matches = sum(
                1 for value in payload if normalize_header(value) in _PDM_HEADER_LABELS
            )
            if matches >= 2:
                continue
        rows.append(payload)
    return rows


def read_pdm_file(
    path: str | Path,
    sheet_name: str | None = None,
) -> PdmReadResult:
    """用 openpyxl 读取 PDM 输入；只取 B:N，不做 TC 表头校验。"""
    file_path = Path(path)
    file_text = str(file_path)
    if file_path.suffix.lower() not in config.INPUT_EXTENSIONS:
        raise InputError(f"只支持 {config.INPUT_EXTENSION_LABEL} 格式：{file_path.name}")
    if not file_path.exists():
        raise InputError(f"文件不存在：{file_path}")
    if not file_path.is_file():
        raise InputError(f"不是文件：{file_path}")

    try:
        workbook = load_workbook(
            filename=file_path,
            data_only=True,
            read_only=False,
            keep_links=False,
        )
    except PermissionError as exc:
        raise InputError(f"无法读取文件（可能被占用）：{file_path}") from exc
    except InvalidFileException as exc:
        raise InputError(f"不是有效的 Excel .xlsx 文件：{file_path}") from exc
    except Exception as exc:
        raise InputError(f"openpyxl 打开失败：{file_path}（{exc}）") from exc

    try:
        issues: list[Issue] = []
        if sheet_name:
            if sheet_name not in workbook.sheetnames:
                raise InputError(f"工作表不存在：{sheet_name}")
            worksheet = workbook[sheet_name]
        else:
            worksheet = workbook.active or workbook.worksheets[0]
            if len(workbook.sheetnames) > 1:
                issues.append(
                    _issue(
                        "info",
                        "MULTI_SHEET",
                        f"工作簿包含多个工作表，使用活动工作表“{worksheet.title}”",
                        file_text,
                        worksheet.title,
                    )
                )
        matrix = [
            list(row)
            for row in worksheet.iter_rows(min_row=1, max_col=14, values_only=True)
        ]
        return PdmReadResult(
            file=file_text,
            sheet=worksheet.title,
            rows=pdm_rows_from_matrix(matrix),
            issues=issues,
            engine="openpyxl",
        )
    finally:
        workbook.close()


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False
