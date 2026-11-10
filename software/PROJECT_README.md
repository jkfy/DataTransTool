# DataTransTool — TC 一键式导入工具

把物料总表（`.xlsx`）按固定模板转换成 **TC 表** 和 **Eplan 表** 的桌面小工具。

- 前端：Qt 6 / C++（Widgets）
- 后端：Python 3 + openpyxl，以独立 CLI 进程运行
- 通信：Qt `QProcess` + JSON 任务文件 + stdout JSONL 进度
- 交付：PyInstaller 打包后端、windeployqt 收集 Qt 运行库，最终生成一个可离线运行的 zip

## 当前状态

已完成可运行版本，包含：

- 多文件选择（列表顺序合并）、拖拽添加；
- 表头自动识别（适配 `inputExample-total.xlsx` 的合并表头布局）；
- 必填项校验（代号、部件编号、备注）；
- `EPLAN` / `新TC` 勾选决定进入哪张输出表；
- TC / Eplan 固定模板输出；
- 文本日志 + `result.json` + 界面进度/错误提示；
- 默认“有错误整批不导出”，可勾选“跳过错误行继续导出”；
- 一键打包脚本，产出离线 zip。

> 本文件是 `software/` 下的项目说明；`backend/`、`ui/`、`templates/`、`packaging/`
> 与本目录的离线包/源码包并列，构建命令可在本目录直接执行。

## 目录结构

```text
DataTransTool/
  backend/                 # Python 后端
    datatrans/
      cli.py               # 命令行入口 / JSONL 协议
      config.py            # 模板、别名、默认值
      headers.py           # 表头识别与别名归一化
      reader.py            # 读取 .xlsx
      validator.py         # 字段解析与校验
      mapper.py            # 记录 -> TC/Eplan 行
      writer.py            # 按模板写出
      report.py            # 文本日志 / result.json
      job.py               # 任务编排
    tests/                 # unittest 测试
    requirements.txt
    requirements-build.txt
  ui/                      # Qt/C++ 前端
    CMakeLists.txt
    main.cpp
    MainWindow.ui          # 控件/布局（Qt Designer 可直接打开）
    app.qrc                # 资源清单，把图标编进 exe
    MainWindow.{h,cpp}     # 槽函数与业务逻辑
    BackendClient.{h,cpp}  # QProcess + JSONL
    tests/                 # QProcess / UI 冒烟测试
  icons/                   # 应用图标
    transIcon.png
  templates/               # 固定输出模板
    output_TC.xlsx
    output_Eplan.xlsx
    input_reference.xlsx   # 输入格式参考，不参与运行
  data/                    # 需求方给的样例
  packaging/
    build.ps1              # 一键构建 + 打包
    README_offline.txt     # 随包说明
  dist/                    # 构建产物（gitignore）
    DataTransTool/
    DataTransTool-0.1.0-win64.zip
```

> 界面控件和布局都在 `ui/MainWindow.ui`，可直接用 Qt Designer 或 Qt Creator 打开修改；
> `MainWindow.cpp` 只保留槽函数、校验和流程控制。CMake 已开启 `AUTOUIC`，改完 `.ui` 重新编译即可。
>
> 应用图标在 `icons/transIcon.png`，通过 `ui/app.qrc` 编进 exe；替换图标后重新编译即可生效。

## 输入格式

按 `data/inputExample-total.xlsx` 的布局解析：

- 只接收 **`.xlsx`**；`.xls`、`.xlsm` 会被拒绝。
- 第 2–3 行是表头（含合并单元格），数据从第 4 行开始。
- 程序不写死行号：会在前 20 行里找命中列名最多的一行，并兼容“组表头 + 子表头”的两行布局。
- 必需的列：`代号`、`部件编号（厂家代号）`、`备注`、`名称`、`EPLAN`、`新TC`。
- 其他列缺失时使用默认值或留空，并在日志里给出警告。

### 勾选规则

- `EPLAN` 为 `√`/`是`/`Y`/`1`/`X` → 进入 Eplan 表；
- `新TC` 为 `√`/`是`/`Y`/`1`/`X` → 进入 TC 表；
- 两列都为空/`×`/`否` → 该行跳过，并写一条警告；
- 两列都勾选 → 同时进入两张表。

## 转换规则

| 物料总表 | TC 表 | Eplan 表 |
|---|---|---|
| 物料类型（空 → 电器外购件） | 采购类型 | |
| 名称 | 名称1 | 名称 |
| 代号 | 类型编码 | 规格型号 |
| 材质（空 → 成品） | 材料 | |
| 部件编号（厂家代号） | 部件编码 | 部件编码 |
| 备注 | 制造商 | 生产商、供应商 |
| 重量（kg） | 重量 | |
| 宽 / 高 / 深（空 → 0） | | 宽度 / 高度 / 深度 |
| | ERP/EPR 编码：**留空** | |
| | 是否合法：**留空** | |
| | | 单位：件 |
| | | 价格：留空 |
| | | 一类产品：`1:电气工程` |
| | | 行业/子行业：`1:电气工程` |
| | | 描述及其他未提到字段：留空 |

> TC 模板表头原文是 `EPR编码`。程序保留模板表头，不修改拼写；该列值留空。

## 输出

输出到界面选择的目录：

- `TC_YYYYMMDD_HHMMSS_xxxxxx.xlsx`
- `EPlan_YYYYMMDD_HHMMSS_xxxxxx.xlsx`
- `import_YYYYMMDD_HHMMSS_xxxxxx.log`（UTF-8 with BOM，Windows 记事本可直接看）
- `result.json`（结构化结果，供界面显示/排错）

## 运行

### 方式一：直接运行离线包

1. 解压 `dist/DataTransTool-0.1.0-win64.zip`。
2. 双击 `DataTransTool.exe`。
3. 添加 `.xlsx` 文件、选择输出目录、点击“开始处理”。

目标机不需要安装 Python、Qt，也不需要联网。详细说明见包内 `README.txt`。

### 方式二：源码开发模式

> 下面命令使用 **PowerShell 语法**（`$env:`、`$PWD`、反引号 `` ` `` 续行），
> 请在 **Windows PowerShell 5.1** 或 **PowerShell 7** 中运行，不要直接粘到 `cmd.exe`。
> 如果必须用 cmd，需要把 `$env:NAME = ...` 改成 `set NAME=...`，把反引号续行改成 `^`。

后端：

```powershell
$env:PYTHONPATH = "$PWD\backend"
python -m datatrans `
  --input "data\inputExample-total.xlsx" `
  --out-dir "output\dev" `
  --templates "templates" `
  --progress
```

前端（开发时让 Qt 调本机 Python，而不是打好的 exe）：

```powershell
$env:DATATRANS_PYTHON = "D:\Program Files\python3,12\python.exe"
$env:DATATRANS_BACKEND_DIR = "$PWD\backend"
$env:Path = "D:\QtCreate\6.10.0\mingw_64\bin;D:\QtCreate\Tools\mingw1310_64\bin;" + $env:Path

cmake -S ui -B build\ui -G "MinGW Makefiles" `
  -DCMAKE_BUILD_TYPE=Release `
  -DCMAKE_PREFIX_PATH="D:\QtCreate\6.10.0\mingw_64" `
  -DCMAKE_C_COMPILER="D:\QtCreate\Tools\mingw1310_64\bin\gcc.exe" `
  -DCMAKE_CXX_COMPILER="D:\QtCreate\Tools\mingw1310_64\bin\g++.exe"

cmake --build build\ui -j 8
.\build\ui\DataTransTool.exe
```

也可以不设 `DATATRANS_PYTHON`，改用 `DATATRANS_CORE` 指向已打包的后端：

```powershell
$env:DATATRANS_CORE = "$PWD\dist\DataTransTool\backend\datatrans-core.exe"
```

## 构建离线包

构建机需要：

- Windows 10/11 64 位；
- Python 3.10+（已验证 3.12.1）；
- Qt 6.5+（已验证 6.10.0 MinGW 64）；
- MinGW 13.1（随 Qt 安装）；
- CMake 3.21+；
- PowerShell 5.1+；
- 构建机可联网（`pip install` 需要；目标内网机器不需要）。

一键构建：

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build.ps1
```

> 打包脚本会先尝试关闭从 `dist\DataTransTool` 启动的 `DataTransTool.exe` / `datatrans-core.exe`，
> 压缩阶段也会自动重试 5 次。若仍提示 `base_library.zip` 被占用，通常是应用没有完全退出，
> 或杀毒软件正在扫描该文件；关闭应用、稍等几秒再执行即可。

常用参数：

```powershell
# 只重新组装/压缩，不重打后端和前端
powershell -ExecutionPolicy Bypass -File packaging\build.ps1 -SkipBackend -SkipUi

# 指定 Qt / MinGW / Python 路径
powershell -ExecutionPolicy Bypass -File packaging\build.ps1 `
  -QtDir "D:\Qt\6.10.0\mingw_64" `
  -MinGWDir "D:\Qt\Tools\mingw1310_64" `
  -Python "python"
```

产物：

- 目录：`dist\DataTransTool\`
- 压缩包：`dist\DataTransTool-0.1.0-win64.zip`（约 33 MB）

## 测试

后端单元测试：

```powershell
$env:PYTHONPATH = "$PWD\backend"
python -m unittest discover -s backend\tests -v
```

Qt 后端通信冒烟测试：

```powershell
cmake -S ui -B build\ui -G "MinGW Makefiles" `
  -DCMAKE_BUILD_TYPE=Release `
  -DCMAKE_PREFIX_PATH="D:\QtCreate\6.10.0\mingw_64" `
  -DDATATRANS_BUILD_SMOKE_TEST=ON

cmake --build build\ui -j 8

$env:DATATRANS_PYTHON = "D:\Program Files\python3,12\python.exe"
$env:DATATRANS_BACKEND_DIR = "$PWD\backend"
.\build\ui\backend_client_smoke.exe `
  "$PWD\data\inputExample-total.xlsx" `
  "$PWD\output\smoke"
```

Qt 界面端到端测试：

```powershell
cmake -S ui -B build\ui -G "MinGW Makefiles" `
  -DCMAKE_BUILD_TYPE=Release `
  -DCMAKE_PREFIX_PATH="D:\QtCreate\6.10.0\mingw_64" `
  -DDATATRANS_BUILD_UI_TEST=ON

cmake --build build\ui -j 8

$env:DATATRANS_SAMPLE = "$PWD\data\inputExample-total.xlsx"
$env:DATATRANS_PYTHON = "D:\Program Files\python3,12\python.exe"
$env:DATATRANS_BACKEND_DIR = "$PWD\backend"
$env:DATATRANS_NO_DIALOGS = "1"
.\build\ui\ui_smoke_test.exe -v1
```

## 离线内网部署说明

**可以。** 最终产物是一个自包含目录/zip：

- `DataTransTool.exe` + Qt DLL + 插件；
- `backend\datatrans-core.exe` + PyInstaller 自带的 Python 运行时；
- `templates\` 固定模板；
- 不需要目标机安装 Python、Qt、Office 或联网。

把 zip 拷进内网，解压到任意目录，双击 `DataTransTool.exe` 即可。注意：

1. **不要只复制 exe**：`backend\`、`templates\`、Qt DLL/插件目录必须保持相对位置。
2. 目标系统：Windows 10/11 64 位；不支持 32 位。
3. 如果内网有杀毒/白名单，未签名的 exe 可能被拦；建议用企业代码签名，或把整个目录加入白名单。
4. 程序不做任何联网请求；日志/结果只写本地输出目录。
5. 首次运行如果 Windows 提示“未知发布者”，属于未签名程序的 SmartScreen 提示，不是程序联网。

## 已知限制 / 约定

- 只支持 `inputExample-total.xlsx` 这一种输入布局；`物料编码申请单数据总模板.xlsx` 的额外列（标准管理员、ERP编码、单位重量、截面积等）未纳入映射。
- 只接收 `.xlsx`，不接收 `.xls` / `.xlsm`。
- TC 的 `EPR编码`、`是否合法` 留空；Eplan 的 `描述`、`价格` 及其他未提到字段留空。
- 多文件按列表顺序合并；重复的“代号 + 部件编号”只写警告，不自动去重。
- 默认“有错误整批不导出”；需要部分导出时勾选“跳过错误行继续导出”。
- 不做界面内预览/编辑；输出后可在 Excel 中检查。
- 表头识别是“按列名”而不是按固定列号，模板列顺序调整通常仍可用；但如果列名大改，需要更新 `backend/datatrans/config.py` 的别名表。

## 第三方组件

- Python（PSF License）
- openpyxl（MIT）
- et_xmlfile（MIT）
- Qt 6（LGPLv3 / GPLv3 / 商业许可，按实际分发方式遵守）
- PyInstaller（GPLv2 + bootloader exception）
- MinGW runtime（GCC Runtime Library Exception）

随包会附带 `THIRD_PARTY_NOTICES.txt`。闭源分发时请按 LGPL 要求保留动态链接、版权声明和可替换库说明；具体请法务确认。

## 故障排查

| 现象 | 处理 |
|---|---|
| 提示“找不到后端程序” | 确认 `backend\datatrans-core.exe` 存在，且未被杀毒隔离 |
| 提示“只支持 .xlsx” | 输入必须是 `.xlsx`；`.xls`/`.xlsm` 请先另存为 `.xlsx` |
| 提示输出文件被占用 | 关闭 Excel 中打开的同名输出文件后重试 |
| 校验失败 | 打开输出目录的 `import_*.log`，按 file/sheet/row/field 定位 |
| 中文日志乱码 | 日志是 UTF-8 with BOM；用记事本/VS Code 打开；不要用旧版 ANSI 编辑器 |
| 程序启动报缺少 DLL | 确认解压完整，Qt DLL 与 `platforms\` 目录未被删 |
