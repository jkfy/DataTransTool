"""通过本机已安装的 Excel（COM）读写工作簿。

适用场景：内网文件被透明加密/DLP 包装，Python 直接读字节会得到
“file is not a zip file”，但本机 Excel 被授权打开并解密。

注意：
- 这不是绕过加密。Excel 仍按当前 Windows 用户的权限打开文件；
  如文件有密码、IRM 权限不足或 DLP 禁止自动化，这里同样会失败。
- 依赖 Windows + 已安装 Excel + pywin32。
"""

from __future__ import annotations

import gc
import os
from pathlib import Path
from typing import Any

from . import config
from .headers import canonical_output_header
from .models import Issue
from .reader import InputError, MatrixSheet, ReadResult, read_grid
from .writer import OutputError


class ExcelComError(InputError):
    """Excel COM 读取失败。"""


def _load_com() -> tuple[Any, Any]:
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception as exc:  # pragma: no cover - 取决于运行环境
        raise ExcelComError(
            "未安装 pywin32，无法使用 Excel 兼容模式。请安装 pywin32 或改用标准模式。"
        ) from exc
    return pythoncom, win32com.client


def _com_error_text(exc: Exception) -> str:
    try:
        args = getattr(exc, "args", ())
        if len(args) >= 2 and isinstance(args[1], tuple) and len(args[1]) >= 3:
            text = args[1][2]
            if text:
                return str(text)
        return str(exc)
    except Exception:
        return repr(exc)


def _new_excel(win32: Any) -> Any:
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    excel.Interactive = False
    # msoAutomationSecurityForceDisable = 3：不执行工作簿里的宏。
    excel.AutomationSecurity = 3
    try:
        excel.AskToUpdateLinks = False
    except Exception:
        pass
    return excel


def _validate_input_path(path: Path) -> None:
    if path.suffix.lower() != config.INPUT_EXTENSION:
        raise InputError(f"只支持 {config.INPUT_EXTENSION} 格式：{path.name}")
    if not path.exists():
        raise InputError(f"文件不存在：{path}")
    if not path.is_file():
        raise InputError(f"不是文件：{path}")


def _is_password_protected_ole(path: Path) -> bool:
    """Office 密码/IRM 保护的 xlsx 常是 OLE 复合文档，而不是 ZIP。

    这种文件让 Excel COM 打开时会弹密码/权限框，自动化会一直等。
    这里直接识别并报错，避免后端卡死；密码/权限本身不能绕过。
    """
    try:
        with path.open("rb") as handle:
            return handle.read(8) == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    except OSError:
        return False


def _to_matrix(raw: Any, rows: int, cols: int) -> list[list[Any]]:
    """把 Excel COM UsedRange.Value 统一成二维列表。"""
    if rows <= 0 or cols <= 0:
        return []
    if rows == 1 and cols == 1:
        return [[raw]]
    if rows == 1:
        if isinstance(raw, (tuple, list)):
            return [list(raw)]
        return [[raw] + [None] * (cols - 1)]
    if not isinstance(raw, (tuple, list)):
        return [[raw] + [None] * (cols - 1)]

    matrix: list[list[Any]] = []
    for row in raw:
        if isinstance(row, (tuple, list)):
            matrix.append(list(row))
        else:
            matrix.append([row])
    return matrix


def read_with_excel(
    path: str | Path,
    sheet_name: str | None = None,
) -> ReadResult:
    """用 Excel COM 打开文件并读取单元格值。

    Excel 会在自己的进程内处理透明加密/解密；Python 只拿到单元格值。
    """
    file_path = Path(path).resolve()
    _validate_input_path(file_path)
    if _is_password_protected_ole(file_path):
        raise ExcelComError(
            f"文件是受密码/IRM 保护的 OLE 格式，Excel 自动化会弹出密码框：{file_path}。"
            "当前版本不绕过密码；请提供密码后另存，或让管理员导出为未加密 .xlsx。"
        )
    pythoncom, win32 = _load_com()
    pythoncom.CoInitialize()

    excel = None
    workbook = None
    worksheet = None
    used = None
    grid = None
    try:
        excel = _new_excel(win32)
        workbook = excel.Workbooks.Open(
            Filename=str(file_path),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
            Notify=False,
        )

        extra_issues: list[Issue] = []
        if sheet_name:
            try:
                worksheet = workbook.Worksheets(sheet_name)
            except Exception as exc:
                raise ExcelComError(f"工作表不存在：{sheet_name}（{_com_error_text(exc)}）") from exc
        else:
            worksheet = workbook.ActiveSheet
            try:
                sheet_count = int(workbook.Worksheets.Count)
            except Exception:
                sheet_count = 1
            if sheet_count > 1:
                extra_issues.append(
                    Issue(
                        level="info",
                        code="MULTI_SHEET",
                        message=f"工作簿包含多个工作表，使用活动工作表“{worksheet.Name}”",
                        file=str(file_path),
                        sheet=str(worksheet.Name),
                    )
                )

        used = worksheet.UsedRange
        first_row = int(used.Row)
        first_col = int(used.Column)
        row_count = int(used.Rows.Count)
        col_count = int(used.Columns.Count)
        matrix = _to_matrix(used.Value, row_count, col_count)
        grid = MatrixSheet(
            matrix,
            first_row=first_row,
            first_col=first_col,
            title=str(worksheet.Name),
        )
        return read_grid(
            grid,
            file_text=str(file_path),
            sheet_title=str(worksheet.Name),
            engine="excel",
            extra_issues=extra_issues,
        )
    except InputError:
        raise
    except Exception as exc:
        raise ExcelComError(
            f"Excel COM 读取失败：{file_path}（{_com_error_text(exc)}）"
        ) from exc
    finally:
        # 先释放工作表/Range 代理，再 Close/Quit，避免 Excel 进程残留或等待。
        worksheet = None
        used = None
        grid = None
        gc.collect()
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        workbook = None
        excel = None
        gc.collect()
        pythoncom.CoUninitialize()


def _header_values(raw: Any, cols: int) -> list[Any]:
    if cols <= 1:
        return [raw]
    if isinstance(raw, (tuple, list)):
        # 单行多列时 Excel COM 常返回 ((v1, v2, ...),)，需要拆一层。
        if len(raw) == 1 and isinstance(raw[0], (tuple, list)):
            return list(raw[0])
        return list(raw)
    return [raw] + [None] * (cols - 1)


def write_table_with_excel(
    *,
    kind: str,
    template_path: str | Path,
    output_path: str | Path,
    rows: Any,
) -> list[str]:
    """用 Excel 打开模板、写入数据并另存为 .xlsx。

    与 writer.write_table 的入参保持一致，便于 job 层切换引擎。
    """
    template = Path(template_path).resolve()
    output = Path(output_path).resolve()
    if not template.exists():
        raise OutputError(f"输出模板不存在：{template}")

    data_rows = [dict(row) for row in rows]
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        try:
            output.unlink()
        except OSError as exc:
            raise OutputError(f"无法覆盖输出文件（可能被占用）：{output}") from exc

    pythoncom, win32 = _load_com()
    pythoncom.CoInitialize()

    excel = None
    workbook = None
    worksheet = None
    used = None
    raw_headers = None
    try:
        excel = _new_excel(win32)
        workbook = excel.Workbooks.Open(
            Filename=str(template),
            UpdateLinks=0,
            ReadOnly=False,
            AddToMru=False,
            Notify=False,
        )
        worksheet = workbook.ActiveSheet

        used = worksheet.UsedRange
        first_col = int(used.Column)
        col_count = int(used.Columns.Count)
        last_row = int(used.Row) + int(used.Rows.Count) - 1
        max_col = first_col + col_count - 1

        raw_headers = worksheet.Range(
            worksheet.Cells(1, first_col), worksheet.Cells(1, max_col)
        ).Value
        headers = _header_values(raw_headers, col_count)

        columns: list[tuple[int, str]] = []
        unknown: list[str] = []
        for offset, header in enumerate(headers):
            col = first_col + offset
            canonical = canonical_output_header(header)
            if canonical:
                columns.append((col, canonical))
            elif header is not None and str(header).strip():
                unknown.append(str(header))

        if not columns:
            raise OutputError(f"无法识别输出模板表头：{template}")

        if last_row >= 2:
            worksheet.Range(
                worksheet.Cells(2, first_col), worksheet.Cells(last_row, max_col)
            ).ClearContents()

        for col, canonical in columns:
            if canonical in config.TEXT_OUTPUT_FIELDS and data_rows:
                worksheet.Range(
                    worksheet.Cells(2, col),
                    worksheet.Cells(1 + len(data_rows), col),
                ).NumberFormat = "@"

        if data_rows:
            matrix: list[list[Any]] = []
            for row in data_rows:
                values: list[Any] = [None] * col_count
                for col, canonical in columns:
                    values[col - first_col] = row.get(canonical)
                matrix.append(values)
            worksheet.Range(
                worksheet.Cells(2, first_col),
                worksheet.Cells(1 + len(data_rows), max_col),
            ).Value = tuple(tuple(row) for row in matrix)

        workbook.SaveAs(str(output), FileFormat=51)  # xlOpenXMLWorkbook
        return unknown
    except OutputError:
        raise
    except Exception as exc:
        raise OutputError(
            f"Excel COM 写出失败：{output}（{_com_error_text(exc)}）"
        ) from exc
    finally:
        # 先释放所有 COM 代理，再 Close/Quit；否则 Excel 偶尔会在 Close 处等待。
        worksheet = None
        used = None
        raw_headers = None
        gc.collect()
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        workbook = None
        excel = None
        gc.collect()
        pythoncom.CoUninitialize()
