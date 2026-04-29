# COMSOL MCP

[中文说明](README.zh-CN.md)

This repository contains an MCP server for controlling an already-open COMSOL Desktop GUI through the COMSOL Java Shell.

The server is designed for GUI-visible edits: Codex pastes COMSOL Java API commands into the Java Shell of the currently open COMSOL Desktop model, then executes them with `Ctrl+Enter`.

## Scope

This repository currently publishes only the GUI Java Shell MCP implementation.

COMSOL can also be controlled through the Python `mph` package in a background workflow, but that implementation is not included here. If you need the background route, keep it as a separate project with its own source ownership and license review.

## Requirements

- Windows
- COMSOL Desktop 6.3 or newer
- COMSOL Desktop and the target model already open
- Java Shell visible

Automatic ribbon clicking is disabled by default because localized COMSOL ribbon automation can trigger COMSOL UI errors.

## Setup

```powershell
.\setup_venv.ps1
```

Optional local settings:

```powershell
Copy-Item .\env.local.example .\env.local
```

## Run

```powershell
.\start_server.ps1
```

Codex global config should point `mcp_servers.comsol_gui` to:

- `E:\AgentCOMSOL\AgentCOMSOL-main\comsol-gui-mcp\.venv\Scripts\python.exe`
- `E:\AgentCOMSOL\AgentCOMSOL-main\comsol-gui-mcp\server.py`

## MCP Tools

- `gui_status()` lists COMSOL GUI processes/windows and whether Java Shell is visible.
- `ensure_java_shell()` finds Java Shell and returns a manual action if it is not visible.
- `execute_java_shell(code, allow_non_model_code=false, timeout_sec=30)` pastes commands and executes them with `Ctrl+Enter`.
- `set_global_parameter(name, value, description=null)` sets a global parameter through Java Shell.
- `get_java_shell_output()` tries to read visible Java Shell text through UI Automation.

## Safe Command Policy

By default, `execute_java_shell` only accepts non-empty executable lines that start with `model.`. Use `allow_non_model_code=true` only when you intentionally need broader Java Shell commands.

## First Validation

1. Open COMSOL Desktop and a test model.
2. Open `Home > Windows > Java Shell`.
3. Call `gui_status()`.
4. Call `set_global_parameter("codex_gui_mcp_probe", "1")`.
5. Confirm the parameter appears in the GUI.
6. Clean up with `execute_java_shell("model.param().remove(\"codex_gui_mcp_probe\");")`.

## Notice

This is not an official COMSOL project. COMSOL product names belong to their respective owner.
