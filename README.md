# PPT Editor

项目已做模块化整理，当前目录说明：

- `src/app.py`：程序启动入口（供打包使用）
- `src/ui.py`：Tkinter 界面逻辑
- `src/services/excel_reader.py`：Excel 读取与结构化
- `src/services/ppt_exporter.py`：PPT 替换和导出服务
- `packaging/`：PyInstaller 配置
- `scripts/`：构建脚本
- `docs/`：文档
- `run.py`：本地启动入口（内部转发到 `src/app.py`）
- `requirements.txt`：依赖列表

本地启动：

```powershell
python .\run.py
```

快速打包：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1
```

发布安装包（推荐）：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_installer.ps1
```
