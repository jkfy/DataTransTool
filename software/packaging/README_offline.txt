# DataTransTool 离线运行说明

## 启动

1. 把整个 `DataTransTool` 文件夹解压到本机任意目录（路径建议不要含特殊字符）。
2. 双击 `DataTransTool.exe`。
3. 添加一个或多个 `.xlsx` 物料总表，选择输出目录，点击“开始处理”。

不需要安装 Python、Qt 或联网。`backend\datatrans-core.exe` 是随包携带的 Python 后端，
`templates\` 是固定输出模板，请勿删除或改名。
`THIRD_PARTY_NOTICES.txt` 是第三方组件许可说明。

## 加密文件

如果内网电脑对 `.xlsx` 做了透明加密，程序会自动切换到本机 Excel 读取/写出：

- 需要本机安装 Microsoft Excel，且 Excel 本身能正常打开该文件；
- 首次调用 Excel 可能较慢（Excel 冷启动/杀毒扫描）；
- 这不是破解：密码、IRM 权限、DLP 策略仍然有效；
- 若提示“受密码/IRM 保护的 OLE 格式”，请提供密码后另存，或让管理员导出为未加密 `.xlsx`；
- 若 Excel 弹出密码/权限对话框，程序无法自动处理，请先手动打开确认权限。

##信息有效性
-为匹配原excel格式，将最后一行含有“注意”的单元格进行特殊处理：如果单元格有“注意”并且该单元格的必填字段都为空，则认为是表格最后一行

## 输出

- `TC_YYYYMMDD_HHMMSS_xxxxxx.xlsx`
- `EPlan_YYYYMMDD_HHMMSS_xxxxxx.xlsx`
- `import_YYYYMMDD_HHMMSS_xxxxxx.log`（校验/处理日志）
- `result.json`（结构化结果，供界面显示）

## 常见问题

- **提示找不到后端程序**：确认 `backend\datatrans-core.exe` 存在，且没有被杀毒软件隔离。
- **提示输出文件被占用**：关闭正在打开同名输出文件的 Excel 后重试。
- **只支持 .xlsx**：输入文件必须是 `.xlsx`；`.xls`/`.xlsm` 不会被接收。
- **校验失败**：查看输出目录下的 `import_*.log`，按文件/行号/字段定位；默认有错误时整批不导出，
  勾选“跳过错误行继续导出”可只导出合法行。

## 系统要求

- Windows 10/11 64 位。
- 普通 `.xlsx` 不需要 Office；透明加密文件需要本机 Microsoft Excel。
- 建议 4 GB 以上内存；数据量通常很小，普通办公电脑即可。
