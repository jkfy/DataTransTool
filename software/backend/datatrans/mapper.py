"""物料总表记录 -> TC / Eplan 输出行。"""

from __future__ import annotations

from typing import Any

from . import config
from .models import MaterialRecord


def map_tc(record: MaterialRecord) -> dict[str, Any]:
    """按设计文档 1.2 生成 TC 行；未提到的字段保持空白。"""
    return {
        "是否合法": None,  # 设计：不填
        "部件编码": record.part_no,
        "名称1": record.name,
        "ERP编码": None,  # 设计：不填（模板表头为 EPR编码，写出时保留模板原文）
        "类型编码": record.code,
        "制造商": record.remark,
        "重量": record.weight,
        "材料": record.material or config.DEFAULT_MATERIAL,
        "采购类型": record.material_type or config.DEFAULT_MATERIAL_TYPE,
    }


def map_eplan(record: MaterialRecord) -> dict[str, Any]:
    """按设计文档 1.2 生成 Eplan 行；未提到的字段保持空白。"""
    return {
        "部件编码": record.part_no,
        "名称": record.name,
        "规格型号": record.code,
        "单位": config.DEFAULT_UNIT,  # 默认：件
        "价格": None,  # 不填
        "生产商": record.remark,
        "供应商": record.remark,
        "描述": None,  # 设计表未要求映射描述
        "一类产品": config.DEFAULT_CATEGORY,  # 固定 1:电气工程
        "行业/子行业": config.DEFAULT_INDUSTRY,  # 固定 1:电气工程
        "宽度": record.width if record.width is not None else 0,
        "高度": record.height if record.height is not None else 0,
        "深度": record.depth if record.depth is not None else 0,
    }
