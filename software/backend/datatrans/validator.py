"""字段解析与业务校验。"""

from __future__ import annotations

import math
import re
from datetime import date, datetime
from typing import Any

from . import config
from .models import Issue, MaterialRecord, RawRow

_NUMBER_RE = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$")


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False


def to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return str(value)
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value).strip()


def to_number(value: Any) -> float | None:
    """把单元格值转成数字；空返回 None；格式不合法抛 ValueError。"""
    if is_blank(value):
        return None
    if isinstance(value, bool):
        raise ValueError("不能是布尔值")
    if isinstance(value, (int, float)):
        number = float(value)
        if math.isnan(number) or math.isinf(number):
            raise ValueError("不是有限数字")
        return number
    text = str(value).strip().replace(",", "")
    if not _NUMBER_RE.match(text):
        raise ValueError("不是有效数字")
    return float(text)


def _issue(
    level: str,
    code: str,
    message: str,
    raw: RawRow,
    field_name: str = "",
    value: Any = None,
) -> Issue:
    return Issue(
        level=level,
        code=code,
        message=message,
        file=raw.file,
        sheet=raw.sheet,
        row=raw.row,
        field=field_name,
        value=value,
    )


def parse_record(raw: RawRow, issues: list[Issue]) -> MaterialRecord:
    """解析一行，按最新设计 1.3 校验。

    必填：部件编号、类型编号、名称1、制造商、供应商。
    重量/宽/高/深：空按 0，负数/非数字报错。
    """
    values = raw.values
    record = MaterialRecord(
        source_file=raw.file,
        source_sheet=raw.sheet,
        source_row=raw.row,
    )

    record.serial = to_text(values.get("序号"))
    record.part_no = to_text(values.get("部件编号"))
    record.erp_code = to_text(values.get("ERP编码"))
    record.code = to_text(values.get("类型编号"))
    record.name = to_text(values.get("名称"))
    record.manufacturer = to_text(values.get("制造商"))
    record.supplier = to_text(values.get("供应商"))
    record.description = to_text(values.get("描述"))
    record.quantity_unit = to_text(values.get("数量单位")) or config.DEFAULT_UNIT

    has_error = False

    for field_label, field_value, raw_value, field_name in (
        ("部件编号", record.part_no, values.get("部件编号"), "部件编号"),
        ("类型编号", record.code, values.get("类型编号"), "类型编号"),
        ("名称1", record.name, values.get("名称"), "名称"),
        ("制造商", record.manufacturer, values.get("制造商"), "制造商"),
        ("供应商", record.supplier, values.get("供应商"), "供应商"),
    ):
        if not field_value:
            has_error = True
            issues.append(
                _issue(
                    "error",
                    "REQUIRED_FIELD_EMPTY",
                    f"{field_label}不能为空",
                    raw,
                    field_name,
                    raw_value,
                )
            )

    for field_label, field_name, attr in (
        ("重量", "重量", "weight"),
        ("宽度", "宽", "width"),
        ("高度", "高", "height"),
        ("深度", "深", "depth"),
    ):
        value = values.get(field_name)
        try:
            number = to_number(value)
        except ValueError as exc:
            has_error = True
            issues.append(
                _issue(
                    "error",
                    "INVALID_NUMBER",
                    f"{field_label}必须是数字（{exc}）",
                    raw,
                    field_name,
                    value,
                )
            )
            continue
        if number is None:
            number = 0.0
        elif number < 0:
            has_error = True
            issues.append(
                _issue(
                    "error",
                    "NEGATIVE_NUMBER",
                    f"{field_label}不能为负数",
                    raw,
                    field_name,
                    value,
                )
            )
        setattr(record, attr, number)

    record.valid = not has_error
    return record
