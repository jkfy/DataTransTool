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


def parse_check(
    value: Any,
    raw: RawRow,
    field_name: str,
    issues: list[Issue],
) -> bool:
    """解析 EPLAN / 新TC 勾选列。空值按未勾选处理。"""
    if is_blank(value):
        return False
    text = to_text(value).upper().strip()
    if text in config.TRUE_VALUES:
        return True
    if text in config.FALSE_VALUES:
        return False
    issues.append(
        _issue(
            "warning",
            "UNKNOWN_CHECK_VALUE",
            f"“{field_name}”的值无法识别为勾选/未勾选，按未勾选处理",
            raw,
            field_name,
            value,
        )
    )
    return False


def parse_record(
    raw: RawRow,
    issues: list[Issue],
    *,
    export_tc: bool = True,
    export_eplan: bool = True,
) -> MaterialRecord:
    values = raw.values
    record = MaterialRecord(
        source_file=raw.file,
        source_sheet=raw.sheet,
        source_row=raw.row,
    )

    record.eplan = parse_check(values.get("EPLAN"), raw, "EPLAN", issues)
    record.new_tc = parse_check(values.get("新TC"), raw, "新TC", issues)

    # 文本字段
    record.material_type = to_text(values.get("物料类型")) or config.DEFAULT_MATERIAL_TYPE
    record.applicant = to_text(values.get("申请人"))
    record.material_code = to_text(values.get("物料编码"))
    record.long_description = to_text(values.get("长描述"))
    record.name = to_text(values.get("名称"))
    record.code = to_text(values.get("代号"))
    record.spec = to_text(values.get("规格"))
    record.material = to_text(values.get("材质")) or config.DEFAULT_MATERIAL
    record.part_no = to_text(values.get("部件编号"))
    record.remark = to_text(values.get("备注"))
    record.voltage = to_text(values.get("额定电压"))
    record.power = to_text(values.get("额定功率"))
    record.description = to_text(values.get("描述"))
    record.outer_diameter = to_text(values.get("外径"))
    record.core_count = to_text(values.get("线芯数"))

    has_error = False

    # 设计文档 1.3 的三个必填项
    for field_name, field_value in (
        ("代号", record.code),
        ("部件编号", record.part_no),
        ("备注", record.remark),
    ):
        if not field_value:
            has_error = True
            issues.append(
                _issue(
                    "error",
                    "REQUIRED_FIELD_EMPTY",
                    f"{field_name}不能为空",
                    raw,
                    field_name,
                    values.get(field_name),
                )
            )

    if not record.name:
        issues.append(
            _issue(
                "warning",
                "NAME_EMPTY",
                "名称为空，输出的名称列也会为空",
                raw,
                "名称",
                values.get("名称"),
            )
        )

    if not record.eplan and not record.new_tc:
        issues.append(
            _issue(
                "warning",
                "NO_EXPORT_TARGET",
                "EPLAN 和 新TC 都未勾选，该行不会进入任何输出表",
                raw,
            )
        )

    # 重量只影响 TC；尺寸只影响 Eplan。未勾选/未启用对应输出时不因该字段阻断。
    if record.new_tc and export_tc:
        weight_value = values.get("重量")
        try:
            record.weight = to_number(weight_value)
        except ValueError as exc:
            has_error = True
            issues.append(
                _issue(
                    "error",
                    "INVALID_NUMBER",
                    f"重量必须是数字（{exc}）",
                    raw,
                    "重量",
                    weight_value,
                )
            )

    if record.eplan and export_eplan:
        for field_name, attr, default_zero in (
            ("宽", "width", True),
            ("高", "height", True),
            ("深", "depth", True),
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
                        f"{field_name}必须是数字（{exc}）",
                        raw,
                        field_name,
                        value,
                    )
                )
                continue
            if number is None:
                number = 0.0 if default_zero else None
            elif number < 0:
                has_error = True
                issues.append(
                    _issue(
                        "error",
                        "NEGATIVE_NUMBER",
                        f"{field_name}不能为负数",
                        raw,
                        field_name,
                        value,
                    )
                )
            setattr(record, attr, number)

    record.valid = not has_error
    return record
