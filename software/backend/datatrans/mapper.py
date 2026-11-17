"""物料记录 -> TC / EPLAN XML 输出行。"""

from __future__ import annotations

from typing import Any

from . import config
from .models import MaterialRecord


def _xml_text(value: str) -> str:
    if not value:
        return ""
    return f"{config.XML_TEXT_PREFIX}{value}{config.XML_TEXT_SUFFIX}"


def _xml_number(value: float | None) -> str:
    if value is None:
        return "0"
    number = float(value)
    if number.is_integer():
        return str(int(number))
    return format(number, ".15g")


def map_tc(record: MaterialRecord) -> dict[str, Any]:
    """按最新设计 1.2 生成 TC 行；未提到的字段留空。"""
    return {
        "是否合法": None,
        "部件编码": record.part_no,
        "名称1": record.name,
        "ERP编码": record.erp_code,
        "类型编码": record.code,
        "制造商": record.manufacturer or record.supplier,
        "重量": record.weight,
        "材料": None,
        "采购类型": None,
    }


def map_xml(record: MaterialRecord) -> dict[str, str]:
    """按 output_Eplan.xml 的 EPLAN 部件属性生成一行。"""
    width = record.width if record.width is not None else 0.0
    height = record.height if record.height is not None else 0.0
    unit = record.quantity_unit or config.DEFAULT_UNIT
    return {
        "P_ARTICLE_PARTTYPE": "1",
        "P_ARTICLE_PARTNR": record.part_no,
        "P_ARTICLE_TYPENR": record.code,
        "P_ARTICLE_ERPNR": record.erp_code,
        "P_ARTICLE_DESCR1": _xml_text(record.name),
        "P_ARTICLE_NOTE": _xml_text(record.description),
        "P_ARTICLE_WEIGHT": _xml_number(record.weight),
        "P_ARTICLE_WIDTH": _xml_number(record.width),
        "P_ARTICLE_HEIGHT": _xml_number(record.height),
        "P_ARTICLE_DEPTH": _xml_number(record.depth),
        "P_ARTICLE_MOUNTINGSPACE": _xml_number(width * height),
        "P_ARTICLE_MANUFACTURER": record.manufacturer,
        "P_ARTICLE_SUPPLIER": record.supplier,
        "P_ARTICLE_QUANTITYUNIT": _xml_text(unit),
        "P_ARTICLE_PRODUCTGROUP": "0",
        "P_ARTICLE_PRODUCTSUBGROUP": "0",
    }
