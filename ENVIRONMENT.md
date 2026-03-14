# 环境安装与运行汇总

本文档用于统一记录本项目的运行环境、安装步骤和常见问题处理方式。

## 1. 推荐环境

- 操作系统：Windows 10/11（已适配 `.bat` 启动脚本）
- Python：3.11（建议 3.10+）
- 包管理器：`pip`

## 2. 快速安装（Windows）

在仓库根目录打开 PowerShell，执行：

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## 3. 启动方式

## 3.1 一键启动（推荐）

双击：

- `启动发票匹配助手.bat`

脚本会：

- 清理占用 `8501` 端口的旧进程
- 启动 Streamlit 服务
- 自动打开浏览器

## 3.2 命令行启动

```powershell
python -m streamlit run streamlit_app.py --server.port 8501
```

## 4. 依赖说明

`requirements.txt` 已覆盖运行所需核心依赖：

- `streamlit`: Web UI 框架
- `pandas` / `openpyxl`: 订单与结果表格读写
- `streamlit-aggrid`: 可编辑表格组件
- `Pillow`: 截图处理
- `pdfplumber` / `pypdf`: 发票 PDF 解析与导出

## 5. 常见问题

## 5.1 启动后页面没更新

- 原因：浏览器缓存或旧进程未释放
- 处理：
  1. 关闭页面
  2. 重新双击 `启动发票匹配助手.bat`
  3. 浏览器执行 `Ctrl+F5`

## 5.2 提示缺少 PDF 解析库

执行：

```powershell
python -m pip install pdfplumber pypdf
```

## 5.3 `streamlit-aggrid` 未安装

执行：

```powershell
python -m pip install streamlit-aggrid
```

