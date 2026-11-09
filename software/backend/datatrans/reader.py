"""读取 .xlsx 输入，定位表头并产出原始行。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from . import config
from .headers import HeaderInfo, describe_header, find_input_header
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


def read_input_file(
    path: str | Path,
    sheet_name: str | None = None,
) -> ReadResult:
    file_path = Path(path)
    file_text = str(file_path)

    if file_path.suffix.lower() != config.INPUT_EXTENSION:
        raise InputError(f"只支持 {config.INPUT_EXTENSION} 格式：{file_path.name}")
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
    except Exception as exc:  # openpyxl 对损坏文件会抛多种异常
        raise InputError(f"打开 Excel 失败：{file_path}（{exc}）") from exc

    result = ReadResult(file=file_text, sheet="", header=None)
    try:
        if sheet_name:
            if sheet_name not in workbook.sheetnames:
                raise InputError(f"工作表不存在：{sheet_name}")
            worksheet = workbook[sheet_name]
        else:
            worksheet = workbook.active or workbook.worksheets[0]
            if len(workbook.sheetnames) > 1:
                result.issues.append(
                    _issue(
                        "info",
                        "MULTI_SHEET",
                        f"工作簿包含多个工作表，使用活动工作表“{worksheet.title}”",
                        file_text,
                        worksheet.title,
                    )
                )

        result.sheet = worksheet.title
        header = find_input_header(worksheet)
        if header is None:
            raise InputError(
                "找不到表头行。请确认输入模板包含“代号、部件编号、备注、名称、EPLAN、新TC”等列。"
            )
        result.header = header

        if header.duplicates:
            result.issues.append(
                _issue(
                    "warning",
                    "DUPLICATE_HEADER",
                    "存在重复表头，已使用第一次出现的列：" + "、".join(header.duplicates),
                    file_text,
                    worksheet.title,
                    header.row,
                )
            )

        missing = header.missing_required
        if missing:
            raise InputError(
                "缺少必需列：" + "、".join(missing) + "。识别到的表头：" + describe_header(header)
            )

        optional_defaults = {
            "物料类型": config.DEFAULT_MATERIAL_TYPE,
            "材质": config.DEFAULT_MATERIAL,
            "重量": "留空",
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
                        worksheet.title,
                        header.row,
                        name,
                    )
                )

        # 逐行读取，不因中间空行提前停止；全空行跳过。
        for row_idx in range(header.data_start_row, worksheet.max_row + 1):
            values: dict[str, Any] = {}
            for canonical, col_idx in header.columns.items():
                values[canonical] = worksheet.cell(row=row_idx, column=col_idx).value
            if all(_is_blank(value) for value in values.values()):
                continue
            result.rows.append(
                RawRow(
                    file=file_text,
                    sheet=worksheet.title,
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
                    worksheet.title,
                    header.data_start_row,
                )
            )
    finally:
        workbook.close()

    return result


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False
