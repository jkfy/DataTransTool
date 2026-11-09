"""按固定模板写出 TC / Eplan 工作簿。"""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from . import config
from .headers import canonical_output_header


class OutputError(Exception):
    """输出文件写入错误。"""


def _fallback_headers(kind: str) -> tuple[str, ...]:
    return config.TC_FALLBACK_HEADERS if kind == "tc" else config.EPLAN_FALLBACK_HEADERS


def _load_template(template_path: Path, kind: str) -> tuple[Any, Any, list[str]]:
    unknown_headers: list[str] = []
    if template_path.exists():
        try:
            workbook = load_workbook(filename=template_path, data_only=False)
        except PermissionError as exc:
            raise OutputError(f"无法读取输出模板（可能被占用）：{template_path}") from exc
        except InvalidFileException as exc:
            raise OutputError(f"输出模板不是有效的 .xlsx 文件：{template_path}") from exc
        except Exception as exc:
            raise OutputError(f"打开输出模板失败：{template_path}（{exc}）") from exc
        worksheet = workbook.active or workbook.worksheets[0]
        for col in range(1, worksheet.max_column + 1):
            raw = worksheet.cell(row=1, column=col).value
            if raw is None or str(raw).strip() == "":
                continue
            if canonical_output_header(raw) is None:
                unknown_headers.append(str(raw))
        return workbook, worksheet, unknown_headers

    # 模板缺失时兜底：至少保证表头和数据结构正确。
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet1"
    for col, header in enumerate(_fallback_headers(kind), start=1):
        worksheet.cell(row=1, column=col, value=header)
    return workbook, worksheet, unknown_headers


def write_table(
    *,
    kind: str,
    template_path: str | Path,
    output_path: str | Path,
    rows: Iterable[dict[str, Any]],
) -> list[str]:
    """把 rows 写入模板，返回模板中未识别的表头（仅用于提示）。"""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook, worksheet, unknown_headers = _load_template(Path(template_path), kind)

    try:
        # 输出模板只保留第 1 行表头：删除说明行/示例数据。
        if worksheet.max_row > 1:
            worksheet.delete_rows(2, worksheet.max_row - 1)

        # 建立“列号 -> 内部规范字段”的映射。
        column_map: dict[int, str] = {}
        for col in range(1, worksheet.max_column + 1):
            raw = worksheet.cell(row=1, column=col).value
            canonical = canonical_output_header(raw)
            if canonical:
                column_map[col] = canonical

        row_index = 2
        for row in rows:
            for col, canonical in column_map.items():
                value = row.get(canonical)
                cell = worksheet.cell(row=row_index, column=col)
                cell.value = value
                if canonical in config.TEXT_OUTPUT_FIELDS:
                    cell.number_format = "@"
            row_index += 1

        tmp_path = output.with_name(output.name + f".{os.getpid()}.tmp.xlsx")
        try:
            workbook.save(tmp_path)
            workbook.close()
            os.replace(tmp_path, output)
        except Exception:
            try:
                workbook.close()
            except Exception:
                pass
            try:
                if tmp_path.exists():
                    tmp_path.unlink()
            except OSError:
                pass
            raise
        return unknown_headers
    except PermissionError as exc:
        raise OutputError(
            f"无法写入输出文件，文件可能正在 Excel 中打开：{output}"
        ) from exc
    except Exception as exc:
        if isinstance(exc, OutputError):
            raise
        raise OutputError(f"写出失败：{output}（{exc}）") from exc


def make_output_paths(
    output_dir: str | Path,
    *,
    timestamp: datetime | None = None,
    job_id: str = "",
) -> dict[str, Path]:
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)
    stamp = (timestamp or datetime.now()).strftime("%Y%m%d_%H%M%S")
    suffix = f"_{job_id[:6]}" if job_id else ""
    base = f"{stamp}{suffix}"

    def unique(name: str) -> Path:
        candidate = directory / name
        if not candidate.exists():
            return candidate
        for index in range(2, 1000):
            candidate = directory / f"{Path(name).stem}_{index}.xlsx"
            if not candidate.exists():
                return candidate
        return directory / f"{Path(name).stem}_{os.getpid()}.xlsx"

    return {
        "tc": unique(f"TC_{base}.xlsx"),
        "eplan": unique(f"EPlan_{base}.xlsx"),
    }
