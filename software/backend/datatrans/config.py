"""应用级配置：路径、模板、表头别名、默认值。"""

from __future__ import annotations

import os
import sys
from pathlib import Path

APP_NAME = "DataTransTool"
VERSION = "0.1.0"

# 输入支持 .xlsx / .xlsm；只读取单元格内容，不执行、不修改宏。
# TC/PDM 输出使用 .xlsx 模板，EPLAN XML 使用 .xml 模板。
INPUT_EXTENSIONS = (".xlsx", ".xlsm")
INPUT_EXTENSION_LABEL = ".xlsx / .xlsm"

DEFAULT_MATERIAL_TYPE = "电器外购件"
DEFAULT_MATERIAL = "成品"
DEFAULT_UNIT = "件"
DEFAULT_CATEGORY = "1:电气工程"
DEFAULT_INDUSTRY = "1:电气工程"

# 业务必填项（设计文档 1.3）
REQUIRED_FIELDS = ("部件编号", "类型编号", "名称", "制造商", "供应商")

# 表格末尾的“注意”说明行（兼容旧模板）：类型编号/部件编号至少缺一个且含“注意”时跳过。
NOTE_ROW_KEYWORDS = ("注意",)
NOTE_ROW_CORE_FIELDS = ("类型编号", "部件编号")

# 固定输入模板中必须存在的表头；缺任何一个都按“格式错误”处理。
REQUIRED_HEADERS = ("部件编号", "类型编号", "名称", "制造商", "供应商")

# 输入表头别名。key 是内部规范名，value 是模板中可能出现的写法。
# 匹配前会统一做全角/半角、空白、换行、尾部括号说明的归一化。
INPUT_ALIASES: dict[str, tuple[str, ...]] = {
    "序号": ("序号",),
    "部件编号": ("部件编号", "部件编码", "厂家代号"),
    "ERP编码": ("ERP编码", "ERP编号", "ERP"),
    "类型编号": ("类型编号", "类型编码", "规格型号", "代号", "型号"),
    "名称": ("名称1", "名称", "部件名称"),
    "制造商": ("制造商", "生产商", "厂家"),
    "供应商": ("供应商", "供货商"),
    "描述": ("描述",),
    "数量单位": ("数量单位", "单位"),
    "重量": ("重量", "重量（kg）", "重量(kg)"),
    "宽": ("宽", "宽度", "尺寸-宽"),
    "高": ("高", "高度", "尺寸-高"),
    "深": ("深", "深度", "尺寸-深"),
}

# 输出表头别名。TC 模板里实际写的是 “EPR编码”，设计文档写 “ERP编码”，
# 这里两种都映射到同一个内部字段；输出时保留模板表头原文，值按输入填写。
OUTPUT_ALIASES: dict[str, tuple[str, ...]] = {
    "是否合法": ("是否合法",),
    "部件编码": ("部件编码",),
    "名称1": ("名称1", "名称"),
    "ERP编码": ("ERP编码", "EPR编码"),
    "类型编码": ("类型编码",),
    "制造商": ("制造商",),
    "重量": ("重量",),
    "材料": ("材料",),
    "采购类型": ("采购类型",),
    "名称": ("名称",),
    "规格型号": ("规格型号",),
    "单位": ("单位",),
    "价格": ("价格",),
    "生产商": ("生产商",),
    "供应商": ("供应商",),
    "描述": ("描述",),
    "一类产品": ("一类产品",),
    "产品组": ("产品组",),
    "子产品组": ("子产品组",),
    "行业/子行业": ("行业/子行业", "行业子行业"),
    "部件分组": ("部件分组",),
    "功能组": ("功能组",),
    "是附件": ("是附件",),
    "安装面": ("安装面",),
    "图形宏": ("图形宏",),
    "宏": ("宏",),
    "宽度": ("宽度", "宽"),
    "高度": ("高度", "高"),
    "深度": ("深度", "深"),
    "左安装间隙": ("左安装间隙",),
    "右安装间隙": ("右安装间隙",),
    "上安装间隙": ("上安装间隙",),
    "下安装间隙": ("下安装间隙",),
    "前安装间隙": ("前安装间隙",),
    "后安装间隙": ("后安装间隙",),
    "功能模板": ("功能模板",),
}

# 输出模板缺失时的兜底表头（正常情况应使用 templates/ 下的固定模板）。
TC_FALLBACK_HEADERS = (
    "是否合法",
    "部件编码",
    "名称1",
    "EPR编码",
    "类型编码",
    "制造商",
    "重量",
    "材料",
    "采购类型",
)
TEMPLATE_FILES = {
    "tc": "output_TC.xlsx",
    "xml": "output_Eplan.xml",
    "pdm": "output_pdm.xlsx",
}
TEMPLATE_FALLBACK_FILES = {
    "tc": "output_Tc.xlsx",
    "xml": "output_Eplan.xml",
    "pdm": "output_pdm.xlsx",
}

# EPLAN 部件管理 XML 中多语言文本的编码格式（与 output_Eplan.xml 一致）。
XML_TEXT_PREFIX = "??_??@"
XML_TEXT_SUFFIX = ";"

# 勾选列的“真值”集合；其他非空值会给出警告并按未勾选处理。
TRUE_VALUES = {"√", "✓", "是", "Y", "YES", "TRUE", "1", "X"}
FALSE_VALUES = {"", "×", "否", "N", "NO", "FALSE", "0", "-", "/"}

# 代码类列，写出时设置为文本格式，避免前导 0 丢失。
TEXT_OUTPUT_FIELDS = {"部件编码", "类型编码", "ERP编码", "规格型号"}


def _source_root() -> Path:
    # backend/datatrans/config.py -> backend/datatrans -> backend -> project root
    return Path(__file__).resolve().parents[2]


def resolve_templates_dir(explicit: str | os.PathLike[str] | None = None) -> Path:
    """按优先级定位模板目录：显式参数 > 环境变量 > 打包目录 > 源码目录。"""
    candidates: list[Path] = []
    if explicit:
        candidates.append(Path(explicit))
    env = os.environ.get("DATATRANS_TEMPLATES")
    if env:
        candidates.append(Path(env))

    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        candidates.append(exe_dir / "templates")
        candidates.append(exe_dir.parent / "templates")
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidates.append(Path(meipass) / "templates")

    root = _source_root()
    candidates.append(root / "templates")
    candidates.append(root / "data")

    for candidate in candidates:
        if (candidate / TEMPLATE_FILES["tc"]).exists() or (
            candidate / TEMPLATE_FALLBACK_FILES["tc"]
        ).exists():
            return candidate
    return candidates[0]


def template_path(kind: str, templates_dir: str | os.PathLike[str] | None = None) -> Path:
    directory = resolve_templates_dir(templates_dir)
    for name in (TEMPLATE_FILES[kind], TEMPLATE_FALLBACK_FILES[kind]):
        path = directory / name
        if path.exists():
            return path
    return directory / TEMPLATE_FILES[kind]
