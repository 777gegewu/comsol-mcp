# COMSOL MCP

[English](README.md)

这个仓库包含一个 MCP 服务器，用于通过 COMSOL Java Shell 控制已打开的 COMSOL Desktop 图形界面。

它面向 GUI 可见编辑：Codex 将 COMSOL Java API 命令粘贴到当前 COMSOL Desktop 模型的 Java Shell 中，并通过 `Ctrl+Enter` 执行。

## 范围

当前仓库只发布 GUI Java Shell MCP 实现。

COMSOL 也可以通过 Python `mph` 包走后台控制路线，但该实现不包含在本仓库中。如果需要后台路线，建议作为独立项目维护，并单独确认源码归属和许可证。

## 环境要求

- Windows
- COMSOL Desktop 6.3 或更新版本
- 已手动打开 COMSOL Desktop 和目标模型
- 已打开 Java Shell 窗口

默认不会自动点击 COMSOL 功能区来打开 Java Shell。不同语言环境下的 COMSOL 功能区自动化不够稳定，可能触发 COMSOL UI 错误。

## 安装

```powershell
.\setup_venv.ps1
```

如需本地配置，可复制示例配置：

```powershell
Copy-Item .\env.local.example .\env.local
```

## 运行

```powershell
.\start_server.ps1
```

Codex 全局配置中的 `mcp_servers.comsol_gui` 应指向：

- `E:\AgentCOMSOL\AgentCOMSOL-main\comsol-gui-mcp\.venv\Scripts\python.exe`
- `E:\AgentCOMSOL\AgentCOMSOL-main\comsol-gui-mcp\server.py`

## MCP 工具

- `gui_status()`：列出 COMSOL GUI 进程和窗口，并检查 Java Shell 是否可见。
- `ensure_java_shell()`：查找 Java Shell；如果不可见，返回需要手动执行的操作。
- `execute_java_shell(code, allow_non_model_code=false, timeout_sec=30)`：将代码粘贴到 Java Shell，并通过 `Ctrl+Enter` 执行。
- `set_global_parameter(name, value, description=null)`：通过 Java Shell 设置全局参数。
- `get_java_shell_output()`：尝试通过 UI Automation 读取 Java Shell 中可见的输出文本。

## 安全策略

默认情况下，`execute_java_shell` 只接受以 `model.` 开头的非空可执行语句。只有在明确需要执行更广泛的 Java Shell 命令时，才应将 `allow_non_model_code` 设置为 `true`。

## 首次验证

1. 打开 COMSOL Desktop 和一个测试模型。
2. 打开 `Home > Windows > Java Shell`。
3. 调用 `gui_status()`，确认 COMSOL 窗口和 Java Shell 可见。
4. 调用 `set_global_parameter("codex_gui_mcp_probe", "1")`。
5. 在 COMSOL GUI 中确认参数已出现。
6. 使用 `execute_java_shell("model.param().remove(\"codex_gui_mcp_probe\");")` 清理测试参数。

## 声明

本项目不是 COMSOL 官方项目。COMSOL 相关产品名称归其权利方所有。
