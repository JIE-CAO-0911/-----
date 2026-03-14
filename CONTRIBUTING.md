# Contributing Guide

感谢你为本项目贡献代码。

## 1. 开发前准备

1. Fork/Clone 仓库
2. 创建虚拟环境并安装依赖（见 [ENVIRONMENT.md](ENVIRONMENT.md)）
3. 新建分支

建议分支命名：

- `feat/<short-name>`
- `fix/<short-name>`
- `docs/<short-name>`

## 2. 提交前自检

提交前请至少完成以下检查：

1. `python -m py_compile streamlit_app.py invoice_matcher.py screenshot_manager.py`
2. 启动应用，手工验证关键流程：
   - 订单预处理
   - 发票匹配
   - 截图归档与导出
3. 文档同步更新（若功能有变化，请更新 README）

## 3. Pull Request 要求

PR 描述建议包含：

1. 背景和目标
2. 修改范围
3. 风险点
4. 验证方式与结果
5. 截图（如涉及 UI）

## 4. 提交信息建议

建议使用简洁前缀：

- `feat: ...`
- `fix: ...`
- `refactor: ...`
- `docs: ...`
- `chore: ...`

## 5. 讨论与提问

如果不确定实现方向，请先提 Issue 说明场景和期望行为，避免重复工作。

