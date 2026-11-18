"""数据模型：原始行、物料记录、校验问题、任务结果。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class RawRow:
    file: str
    sheet: str
    row: int
    values: dict[str, Any]


@dataclass(slots=True)
class Issue:
    level: str  # error / warning / info
    code: str
    message: str
    file: str = ""
    sheet: str = ""
    row: int | None = None
    field: str = ""
    value: Any = None

    def key(self) -> tuple:
        return (self.level, self.code, self.file, self.sheet, self.row, self.field)

    def to_dict(self) -> dict[str, Any]:
        return {
            "level": self.level,
            "code": self.code,
            "message": self.message,
            "file": self.file,
            "sheet": self.sheet,
            "row": self.row,
            "field": self.field,
            "value": _json_value(self.value),
        }


@dataclass(slots=True)
class MaterialRecord:
    # 规范字段（对应 input_reference.xlsx 的列）
    serial: str = ""
    part_no: str = ""
    erp_code: str = ""
    code: str = ""
    name: str = ""
    manufacturer: str = ""
    supplier: str = ""
    description: str = ""
    quantity_unit: str = ""
    weight: float | None = None
    width: float | None = None
    height: float | None = None
    depth: float | None = None

    # 来源信息
    source_file: str = ""
    source_sheet: str = ""
    source_row: int = 0

    valid: bool = True

    def to_dict(self) -> dict[str, Any]:
        return {
            "serial": self.serial,
            "part_no": self.part_no,
            "erp_code": self.erp_code,
            "code": self.code,
            "name": self.name,
            "manufacturer": self.manufacturer,
            "supplier": self.supplier,
            "description": self.description,
            "quantity_unit": self.quantity_unit,
            "weight": self.weight,
            "width": self.width,
            "height": self.height,
            "depth": self.depth,
            "source_file": self.source_file,
            "source_sheet": self.source_sheet,
            "source_row": self.source_row,
            "valid": self.valid,
        }


@dataclass(slots=True)
class FileSummary:
    file: str
    sheet: str = ""
    rows_read: int = 0
    rows_valid: int = 0
    rows_invalid: int = 0
    error: str = ""
    engine: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "file": self.file,
            "sheet": self.sheet,
            "rows_read": self.rows_read,
            "rows_valid": self.rows_valid,
            "rows_invalid": self.rows_invalid,
            "error": self.error,
            "engine": self.engine,
        }


@dataclass(slots=True)
class JobOptions:
    output_dir: str
    export_tc: bool = True
    export_xml: bool = False
    export_pdm: bool = False
    on_invalid: str = "abort"  # abort / skip
    engine: str = "auto"  # auto / openpyxl / excel
    sheet_name: str | None = None
    templates_dir: str | None = None

    @property
    def allow_partial(self) -> bool:
        return self.on_invalid == "skip"


@dataclass(slots=True)
class JobResult:
    job_id: str
    status: str  # success / partial / failed
    ok: bool
    outputs: dict[str, str | None] = field(default_factory=dict)
    stats: dict[str, Any] = field(default_factory=dict)
    issues: list[Issue] = field(default_factory=list)
    log_path: str = ""
    result_path: str = ""
    message: str = ""
    files: list[FileSummary] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": "result",
            "job_id": self.job_id,
            "ok": self.ok,
            "status": self.status,
            "message": self.message,
            "outputs": self.outputs,
            "stats": self.stats,
            "issues": [issue.to_dict() for issue in self.issues],
            "log_path": self.log_path,
            "result_path": self.result_path,
            "files": [item.to_dict() for item in self.files],
        }


def _json_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)
