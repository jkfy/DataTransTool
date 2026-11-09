"""表头归一化、别名解析、输入/输出表头定位。"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from . import config


_TRAILING_PAREN = re.compile(r"\([^()]*\)\s*$")


def normalize_header(value: Any) -> str:
    """把表头统一成可比较的形式。

    处理：全角转半角、去换行/空白、去尾部括号说明、统一大写。
    例如：
      部件编号（厂家代号） -> 部件编号
      重量（kg）           -> 重量
      额定电压（或工作电压范围） -> 额定电压
      EPlan（√）           -> EPLAN
    """
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", "", text)
    # 去掉尾部括号说明；循环处理 "xxx(a)(b)" 这类情况。
    while True:
        new_text = _TRAILING_PAREN.sub("", text)
        if new_text == text:
            break
        text = new_text
    text = text.strip(":：")
    return text.upper()


def _build_alias_map(aliases: dict[str, tuple[str, ...]]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for canonical, names in aliases.items():
        mapping[normalize_header(canonical)] = canonical
        for name in names:
            mapping[normalize_header(name)] = canonical
    return mapping


_INPUT_ALIAS_MAP = _build_alias_map(config.INPUT_ALIASES)
_OUTPUT_ALIAS_MAP = _build_alias_map(config.OUTPUT_ALIASES)


def canonical_input_header(value: Any) -> str | None:
    return _INPUT_ALIAS_MAP.get(normalize_header(value))


def canonical_output_header(value: Any) -> str | None:
    return _OUTPUT_ALIAS_MAP.get(normalize_header(value))


def _header_cell_value(ws: Worksheet, row: int, col: int, allow_above: bool = True) -> Any:
    value = ws.cell(row=row, column=col).value
    if value is None and allow_above and row > 1:
        # 合并表头：值通常在合并区域左上角（上一行）。
        value = ws.cell(row=row - 1, column=col).value
    return value


@dataclass(slots=True)
class HeaderInfo:
    row: int
    columns: dict[str, int]
    raw_headers: dict[int, str] = field(default_factory=dict)
    data_start_row: int = 0
    duplicates: list[str] = field(default_factory=list)

    @property
    def missing_required(self) -> list[str]:
        return [name for name in config.REQUIRED_HEADERS if name not in self.columns]


def find_input_header(ws: Worksheet, max_scan_rows: int = 20) -> HeaderInfo | None:
    """在前若干行里找最像表头的一行。

    支持两类模板：
    - inputExample-total.xlsx：第 3 行是列名，Q-U 的列名在合并的第 2 行；
    - 物料编码申请单数据总模板.xlsx：第 4 行是组表头，第 5 行是“宽/高/深”等子表头。
    """
    best: tuple[tuple[int, int, int], HeaderInfo] | None = None
    limit = min(ws.max_row, max_scan_rows)

    for row in range(1, limit + 1):
        columns: dict[str, int] = {}
        raw_headers: dict[int, str] = {}
        duplicates: list[str] = []
        for col in range(1, ws.max_column + 1):
            raw = _header_cell_value(ws, row, col, allow_above=True)
            if raw is None:
                continue
            text = str(raw).strip()
            if text:
                raw_headers[col] = text
            canonical = canonical_input_header(raw)
            if not canonical:
                continue
            if canonical in columns:
                duplicates.append(canonical)
                continue
            columns[canonical] = col

        required_hits = sum(1 for name in config.REQUIRED_FIELDS if name in columns)
        score = len(columns)
        if score < 3:
            continue
        candidate = HeaderInfo(
            row=row,
            columns=columns,
            raw_headers=raw_headers,
            data_start_row=row + 1,
            duplicates=duplicates,
        )
        key = (required_hits, score, row)
        if best is None or key > best[0]:
            best = (key, candidate)

    if best is None:
        return None

    header = best[1]

    # 补充“组表头 + 子表头”两行模板：例如第 4 行是“尺寸(mm)”，
    # 第 5 行才是“宽/高/深”。只对主表头没识别出来的列向下看一行。
    used_subheader = False
    if header.row < ws.max_row:
        for col in range(1, ws.max_column + 1):
            if any(existing == col for existing in header.columns.values()):
                continue
            raw = ws.cell(row=header.row + 1, column=col).value
            canonical = canonical_input_header(raw)
            if canonical and canonical not in header.columns:
                header.columns[canonical] = col
                if raw is not None:
                    header.raw_headers[col] = str(raw).strip()
                used_subheader = True
    if used_subheader:
        header.data_start_row = header.row + 2

    return header


def describe_header(header: HeaderInfo) -> str:
    """把识别到的表头整理成一行，便于写日志。"""
    parts = []
    for canonical, col in sorted(header.columns.items(), key=lambda item: item[1]):
        parts.append(f"{canonical}={col}")
    return ", ".join(parts)
