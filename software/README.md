# software 目录说明

本目录是 DataTransTool 的交付目录，**代码文件直接放在这里**，不再套一层 `src/`。

> 说明：本轮已把 Qt 控件改为 `ui/MainWindow.ui`（Qt Designer 可直接打开），
> 但按你的要求**暂未重新打包**；现有 `DataTransTool-0.1.0-win64.zip` 是上一版功能等价的运行包。
> 需要更新运行包时，在本目录执行 `packaging\build.ps1` 即可。

## 目录内容

| 路径 | 说明 |
|---|---|
| `DataTransTool-0.1.0-win64.zip` | **离线运行包**。解压后双击 `DataTransTool.exe`，目标机无需 Python/Qt/联网 |
| `DataTransTool-0.1.0-source.zip` | 源码压缩包，便于归档 |
| `backend/` | Python 后端（读取、校验、映射、写 Excel、CLI 协议） |
| `ui/` | Qt/C++ 前端；`MainWindow.ui` 是可视化界面文件，`MainWindow.cpp` 是逻辑 |
| `templates/` | 固定输出模板 |
| `packaging/` | `build.ps1` 一键构建/打包脚本、离线说明、第三方许可 |
| `data/` | 需求方给的输入/输出样例 |
| `PROJECT_README.md` | 项目详细说明（构建、测试、规则、排错） |
| `README.md` | 本文件 |

## 快速运行

```text
解压 DataTransTool-0.1.0-win64.zip
双击 DataTransTool.exe
添加 .xlsx 物料总表
选择输出目录
点击“开始处理”
```

输出：`TC_*.xlsx`、`EPlan_*.xlsx`、`import_*.log`、`result.json`。

## 在 software 目录内重新构建

前提：Windows 10/11 64 位、Python 3.10+、Qt 6.5+（MinGW）、CMake 3.21+，构建机可联网安装 pip 依赖。

```powershell
cd software
powershell -ExecutionPolicy Bypass -File packaging\build.ps1
```

产物会写到：

- `software\dist\DataTransTool\`
- `software\dist\DataTransTool-0.1.0-win64.zip`

不会覆盖本目录已有的离线包；确认新包没问题后，再手动替换 `DataTransTool-0.1.0-win64.zip`。

## 系统要求

- Windows 10/11 64 位；
- 运行包不需要安装 Python、Qt、Office，也不需要联网；
- 未签名 exe 可能被内网杀毒/SmartScreen 拦截，建议代码签名或目录白名单。
