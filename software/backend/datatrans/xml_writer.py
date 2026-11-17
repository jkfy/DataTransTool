"""按 output_Eplan.xml 模板导出 EPLAN 部件管理 XML。"""

from __future__ import annotations

import os
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Iterable

from .mapper import map_xml
from .models import MaterialRecord
from .writer import OutputError


def write_parts_xml(
    template_path: str | Path,
    output_path: str | Path,
    records: Iterable[MaterialRecord],
) -> None:
    """保留模板根节点属性，替换全部 part 子节点。"""
    template = Path(template_path)
    output = Path(output_path)
    if not template.exists():
        raise OutputError(f"XML 模板不存在：{template}")

    try:
        tree = ET.parse(template)
        root = tree.getroot()
    except Exception as exc:
        raise OutputError(f"读取 XML 模板失败：{template}（{exc}）") from exc

    if root.tag != "partsmanagement":
        raise OutputError(f"XML 模板根节点不是 partsmanagement：{template}")

    for child in list(root):
        root.remove(child)

    # records 可能是一次性迭代器；统一成 list 再遍历/计数。
    data = list(records)
    root.set("count", str(len(data)))
    for record in data:
        ET.SubElement(root, "part", map_xml(record))

    ET.indent(tree, space="  ")
    output.parent.mkdir(parents=True, exist_ok=True)
    tmp = output.with_name(output.name + ".tmp")
    try:
        tree.write(tmp, encoding="utf-8", xml_declaration=True)
        os.replace(tmp, output)
    except PermissionError as exc:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
        raise OutputError(f"无法写入 XML（文件可能被占用）：{output}") from exc
    except Exception as exc:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
        raise OutputError(f"写出 XML 失败：{output}（{exc}）") from exc
