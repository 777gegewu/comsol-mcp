import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Any

try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    load_dotenv = None

try:
    from mcp.server.fastmcp import FastMCP
except ModuleNotFoundError:
    FastMCP = None

try:
    import pyperclip
except ModuleNotFoundError:
    pyperclip = None

try:
    import psutil
except ModuleNotFoundError:
    psutil = None

try:
    from PIL import Image
except ModuleNotFoundError:
    Image = None

try:
    import pythoncom
    import win32gui
    import win32process
except ModuleNotFoundError:
    pythoncom = None
    win32gui = None
    win32process = None

try:
    from pywinauto import Application, Desktop, keyboard, timings
except ModuleNotFoundError:
    Application = None
    Desktop = None
    keyboard = None
    timings = None


if getattr(sys, "frozen", False):
    SERVER_DIR = Path(sys.executable).resolve().parent
else:
    SERVER_DIR = Path(__file__).resolve().parent

ENV_FILE = SERVER_DIR / "env.local"
if load_dotenv is not None:
    load_dotenv(ENV_FILE)

COMSOL_PROCESS_NAMES = tuple(
    name.strip().lower()
    for name in os.getenv("COMSOL_GUI_PROCESS_NAMES", "ComsolUI.exe,comsol.exe").split(",")
    if name.strip()
)
WINDOW_TITLE_PATTERN = os.getenv("COMSOL_GUI_WINDOW_TITLE_PATTERN", "COMSOL")
JAVA_SHELL_TITLE_PATTERN = os.getenv("COMSOL_JAVA_SHELL_TITLE_PATTERN", "Java Shell")
DEFAULT_TIMEOUT = float(os.getenv("COMSOL_GUI_TIMEOUT_SEC", "10"))
DESCENDANT_LIMIT = int(os.getenv("COMSOL_GUI_DESCENDANT_LIMIT", "1200"))
SHELL_SCAN_TIMEOUT = float(os.getenv("COMSOL_GUI_SHELL_SCAN_TIMEOUT_SEC", "8"))
AUTO_OPEN_SHELL = os.getenv("COMSOL_GUI_AUTO_OPEN_SHELL", "0").strip().lower() in {"1", "true", "yes"}
WINDOW_LIST_LIMIT = int(os.getenv("COMSOL_GUI_WINDOW_LIST_LIMIT", "20"))
SCREENSHOT_DIR = Path(os.getenv("COMSOL_GUI_SCREENSHOT_DIR", str(SERVER_DIR / "screenshots")))


def _missing_dependencies() -> list[str]:
    missing = []
    if load_dotenv is None:
        missing.append("python-dotenv")
    if FastMCP is None:
        missing.append("mcp")
    if Application is None or Desktop is None or keyboard is None or timings is None:
        missing.append("pywinauto")
    if pyperclip is None:
        missing.append("pyperclip")
    if psutil is None:
        missing.append("psutil")
    if Image is None:
        missing.append("Pillow")
    if win32gui is None or win32process is None or pythoncom is None:
        missing.append("pywin32")
    return missing


def _runtime_error_message() -> str:
    missing = _missing_dependencies()
    if not missing:
        return ""
    return (
        "Missing Python dependencies: "
        + ", ".join(missing)
        + f". Create `{SERVER_DIR}\\.venv` and install `{SERVER_DIR}\\requirements.txt`."
    )


def _ensure_runtime_ready() -> None:
    message = _runtime_error_message()
    if message:
        raise RuntimeError(message)


class _MissingFastMCP:
    def tool(self):
        def decorator(func):
            return func

        return decorator

    def run(self) -> None:
        raise RuntimeError(_runtime_error_message())


mcp = FastMCP("COMSOL GUI MCP") if FastMCP is not None else _MissingFastMCP()


def _matches(pattern: str, text: str) -> bool:
    if not pattern:
        return True
    try:
        return re.search(pattern, text or "", re.IGNORECASE) is not None
    except re.error:
        return pattern.lower() in (text or "").lower()


def _enum_visible_windows() -> list[dict[str, Any]]:
    _ensure_runtime_ready()
    windows: list[dict[str, Any]] = []

    def callback(hwnd, _extra):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd) or ""
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:
            pid = None
        windows.append({"hwnd": hwnd, "pid": pid, "title": title})
        return True

    try:
        win32gui.EnumWindows(callback, None)
    except Exception as exc:
        raise RuntimeError(
            "Windows desktop window enumeration failed. GUI automation must run with permission "
            f"to inspect the interactive desktop. Original error: {exc}"
        ) from exc
    return windows


def _process_name(pid: int | None) -> str | None:
    if pid is None or psutil is None:
        return None
    try:
        return psutil.Process(pid).name()
    except Exception:
        return None


def _find_comsol_processes() -> list[dict[str, Any]]:
    _ensure_runtime_ready()
    found: list[dict[str, Any]] = []
    for proc in psutil.process_iter(["pid", "name", "exe"]):
        try:
            name = proc.info.get("name") or ""
            if name.lower() in COMSOL_PROCESS_NAMES:
                found.append(
                    {
                        "pid": proc.info.get("pid"),
                        "name": name,
                        "path": proc.info.get("exe"),
                    }
                )
        except Exception:
            continue
    return found


def _is_probable_comsol_window(window: dict[str, Any], comsol_pids: set[int]) -> bool:
    title = window.get("title") or ""
    pid = window.get("pid")
    if pid in comsol_pids:
        return True
    if title and _matches(WINDOW_TITLE_PATTERN, title):
        return True
    return False


def _find_comsol_windows() -> list[dict[str, Any]]:
    processes = _find_comsol_processes()
    pids = {p["pid"] for p in processes if p.get("pid") is not None}
    windows = []
    for window in _enum_visible_windows():
        if _is_probable_comsol_window(window, pids):
            item = dict(window)
            item["process_name"] = _process_name(item.get("pid"))
            windows.append(item)
    windows.sort(key=_comsol_window_sort_key)
    if len(windows) <= WINDOW_LIST_LIMIT:
        return windows
    main_windows = [window for window in windows if _is_main_comsol_window(window)]
    shell_or_named = [
        window
        for window in windows
        if window not in main_windows
        and window.get("title")
        and _matches(JAVA_SHELL_TITLE_PATTERN, window.get("title") or "")
    ]
    remainder = [window for window in windows if window not in main_windows and window not in shell_or_named]
    return (main_windows + shell_or_named + remainder)[:WINDOW_LIST_LIMIT]


def _comsol_window_sort_key(window: dict[str, Any]) -> tuple[int, str]:
    title = window.get("title") or ""
    if "COMSOL Multiphysics" in title:
        return (0, title)
    if title and title not in {"ActiproWindowChromeShadow", "错误", "Error"}:
        return (1, title)
    return (2, title)


def _is_main_comsol_window(window: dict[str, Any]) -> bool:
    title = window.get("title") or ""
    return "COMSOL Multiphysics" in title


def _desktop():
    _ensure_runtime_ready()
    pythoncom.CoInitialize()
    timings.Timings.after_clickinput_wait = 0.2
    timings.Timings.window_find_timeout = DEFAULT_TIMEOUT
    return Desktop(backend="uia")


def _control_text(control) -> str:
    try:
        text = control.window_text()
        if text:
            return str(text)
    except Exception:
        pass
    try:
        text = control.element_info.name
        if text:
            return str(text)
    except Exception:
        pass
    return ""


def _walk_children_limited(root, limit: int = DESCENDANT_LIMIT, timeout_sec: float = SHELL_SCAN_TIMEOUT):
    deadline = time.monotonic() + timeout_sec
    stack = [root]
    seen = 0
    while stack and seen < limit and time.monotonic() < deadline:
        control = stack.pop()
        seen += 1
        yield control
        try:
            children = control.children()
        except Exception:
            children = []
        stack.extend(reversed(children))


def _has_edit_like_descendant(root) -> bool:
    for control in _walk_children_limited(root, limit=250, timeout_sec=2):
        try:
            ctype = control.element_info.control_type
        except Exception:
            ctype = ""
        if ctype in {"Edit", "Document"}:
            return True
    return False


def _nearest_shell_container(control):
    current = control
    fallback = control
    for _ in range(8):
        if current is None:
            break
        fallback = current
        if _has_edit_like_descendant(current):
            return current
        try:
            current = current.parent()
        except Exception:
            break
    return fallback


def _comsol_uia_roots() -> list[Any]:
    roots = []
    for window in _find_comsol_windows():
        if not _is_main_comsol_window(window):
            continue
        try:
            roots.append(_connect_top_window(window.get("hwnd")))
        except Exception:
            continue
    return roots


def _find_java_shell_window():
    desktop = _desktop()
    for window in desktop.windows():
        title = _control_text(window)
        if title and _matches(JAVA_SHELL_TITLE_PATTERN, title):
            return window

    # COMSOL often docks Java Shell as a pane inside the main window.
    for top in _comsol_uia_roots():
        for child in _walk_children_limited(top):
            title = _control_text(child)
            if title and _matches(JAVA_SHELL_TITLE_PATTERN, title):
                return _nearest_shell_container(child)
    return None


def _connect_top_window(hwnd: int | None = None):
    _ensure_runtime_ready()
    pythoncom.CoInitialize()
    if hwnd:
        app = Application(backend="uia").connect(handle=hwnd, timeout=DEFAULT_TIMEOUT)
        return app.window(handle=hwnd)

    windows = _find_comsol_windows()
    if not windows:
        raise RuntimeError(
            "No visible COMSOL Desktop window was found. Open COMSOL Desktop and the target model first."
        )

    selected = next((window for window in windows if _is_main_comsol_window(window)), windows[0])
    try:
        app = Application(backend="uia").connect(process=selected["pid"], timeout=DEFAULT_TIMEOUT)
        return app.top_window()
    except Exception:
        app = Application(backend="uia").connect(handle=selected["hwnd"], timeout=DEFAULT_TIMEOUT)
        return app.window(handle=selected["hwnd"])


def _focus_control(control) -> None:
    try:
        control.set_focus()
        time.sleep(0.2)
        return
    except Exception:
        pass
    try:
        control.click_input()
    except Exception:
        pass
    time.sleep(0.2)


def _safe_filename_part(value: str) -> str:
    clean = re.sub(r"[^A-Za-z0-9_.-]+", "_", value.strip())
    return clean[:80] or "comsol"


def _screenshot_path(prefix: str) -> Path:
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    return SCREENSHOT_DIR / f"{_safe_filename_part(prefix)}_{stamp}.png"


def _capture_control(control, prefix: str) -> dict[str, Any]:
    _ensure_runtime_ready()
    _focus_control(control)
    try:
        image = control.capture_as_image()
    except Exception as exc:
        raise RuntimeError(f"Could not capture COMSOL GUI control: {exc}") from exc
    output_path = _screenshot_path(prefix)
    image.save(output_path)
    rect = control.rectangle()
    return {
        "ok": True,
        "path": str(output_path),
        "rect": {"left": rect.left, "top": rect.top, "right": rect.right, "bottom": rect.bottom},
        "width": rect.width(),
        "height": rect.height(),
    }


def _find_graphics_control():
    top = _connect_top_window()
    best = None
    best_area = 0
    for control in _walk_children_limited(top, limit=1800, timeout_sec=12):
        name = _control_text(control)
        try:
            ctype = control.element_info.control_type
            rect = control.rectangle()
        except Exception:
            continue
        if name and _matches(r"^(图形|Graphics)$", name):
            return control
        if ctype in {"Pane", "Custom"}:
            area = max(rect.width(), 0) * max(rect.height(), 0)
            if area > best_area and rect.width() > 300 and rect.height() > 250:
                best = control
                best_area = area
    return best


def _find_named_control(root, pattern: str):
    for control in _walk_children_limited(root, limit=1400, timeout_sec=10):
        if _matches(pattern, _control_text(control)):
            return control
    return None


def _try_click_named(root, pattern: str) -> bool:
    control = _find_named_control(root, pattern)
    if control is None:
        return False
    try:
        control.click_input()
        time.sleep(0.4)
        return True
    except Exception:
        return False


def _open_java_shell_from_gui() -> bool:
    """Best-effort automation for Home > Windows > Java Shell."""
    if not AUTO_OPEN_SHELL:
        return False

    top = _connect_top_window()
    _focus_control(top)

    click_sequences = [
        ["^Home$", "^Windows$", "Java Shell"],
        ["Home", "Windows", "Java Shell"],
        ["Windows", "Java Shell"],
    ]
    for sequence in click_sequences:
        current_root = top
        clicked_any = False
        for pattern in sequence:
            if _try_click_named(current_root, pattern):
                clicked_any = True
                time.sleep(0.5)
                current_root = _desktop()
        if clicked_any and _find_java_shell_window() is not None:
            return True

    # Keyboard fallback. This is layout-dependent, so failure is expected on some builds.
    for sequence in ("{VK_MENU}h", "{VK_MENU}w"):
        try:
            keyboard.send_keys(sequence, pause=0.05)
            time.sleep(0.6)
            if _find_java_shell_window() is not None:
                return True
        except Exception:
            continue
    return False


def _find_edit_like_descendants(root) -> list[Any]:
    controls = []
    for control in _walk_children_limited(root, limit=500, timeout_sec=3):
        try:
            ctype = control.element_info.control_type
        except Exception:
            ctype = ""
        try:
            name = control.element_info.name or ""
        except Exception:
            name = ""
        if ctype in {"Edit", "Document"} or "input" in name.lower() or "command" in name.lower():
            controls.append(control)
    return controls


def _find_shell_input(shell):
    edit_controls = _find_edit_like_descendants(shell)
    if not edit_controls:
        raise RuntimeError(
            "Java Shell was found, but no editable input control was detected. "
            "Click the Java Shell input area manually and retry."
        )
    return edit_controls[-1]


def _validate_java_shell_code(code: str, allow_non_model_code: bool) -> list[str]:
    if not code or not code.strip():
        raise ValueError("Java Shell code is empty.")
    lines = [line.strip() for line in code.splitlines()]
    executable_lines = [line for line in lines if line and not line.startswith("//")]
    if not executable_lines:
        raise ValueError("Java Shell code contains no executable lines.")
    if allow_non_model_code:
        return executable_lines

    bad_lines = []
    for line in executable_lines:
        if line in {";", "{", "}"}:
            continue
        if not line.startswith("model."):
            bad_lines.append(line)
    if bad_lines:
        raise ValueError(
            "Rejected Java Shell code because allow_non_model_code=false and these lines do not start with `model.`: "
            + "; ".join(bad_lines[:5])
        )
    return executable_lines


def _execute_in_shell(code: str, timeout_sec: float) -> dict[str, Any]:
    shell = _find_java_shell_window()
    if shell is None:
        opened = _open_java_shell_from_gui()
        shell = _find_java_shell_window() if opened else None
    if shell is None:
        raise RuntimeError(
            "Could not find or open COMSOL Java Shell. In COMSOL Desktop, open Home > Windows > Java Shell, "
            "click the Java Shell input area once, then retry."
        )

    _focus_control(shell)
    input_control = _find_shell_input(shell)
    _focus_control(input_control)

    previous_clipboard = None
    try:
        previous_clipboard = pyperclip.paste()
    except Exception:
        pass

    try:
        pyperclip.copy(code)
        keyboard.send_keys("^a", pause=0.05)
        keyboard.send_keys("^v", pause=0.05)
        time.sleep(0.2)
        keyboard.send_keys("^{ENTER}", pause=0.05)
        time.sleep(min(max(float(timeout_sec), 0.5), 5.0))
    finally:
        if previous_clipboard is not None:
            try:
                pyperclip.copy(previous_clipboard)
            except Exception:
                pass

    return {
        "submitted": True,
        "java_shell_found": True,
        "note": "Command was submitted to COMSOL Java Shell. Check COMSOL GUI/Java Shell output for runtime errors.",
    }


def _java_string(value: str) -> str:
    return json.dumps(value, ensure_ascii=False)


@mcp.tool()
def gui_status() -> dict[str, Any]:
    """List COMSOL GUI processes/windows and report whether Java Shell is visible."""
    _ensure_runtime_ready()
    processes = _find_comsol_processes()
    window_error = None
    shell_error = None
    try:
        windows = _find_comsol_windows()
    except Exception as exc:
        windows = []
        window_error = str(exc)
    try:
        shell = _find_java_shell_window() if window_error is None else None
    except Exception as exc:
        shell = None
        shell_error = str(exc)
    return {
        "ok": bool(processes or windows),
        "comsol_processes": processes,
        "comsol_windows": windows,
        "java_shell_detected": shell is not None,
        "java_shell_title": _control_text(shell) if shell is not None else None,
        "window_error": window_error,
        "java_shell_error": shell_error,
        "env": {
            "COMSOL_GUI_PROCESS_NAMES": ",".join(COMSOL_PROCESS_NAMES),
            "COMSOL_GUI_WINDOW_TITLE_PATTERN": WINDOW_TITLE_PATTERN,
            "COMSOL_JAVA_SHELL_TITLE_PATTERN": JAVA_SHELL_TITLE_PATTERN,
            "COMSOL_GUI_AUTO_OPEN_SHELL": AUTO_OPEN_SHELL,
            "COMSOL_GUI_WINDOW_LIST_LIMIT": WINDOW_LIST_LIMIT,
        },
    }


@mcp.tool()
def ensure_java_shell() -> dict[str, Any]:
    """Find or open COMSOL Java Shell in the running Desktop GUI."""
    _ensure_runtime_ready()
    shell = _find_java_shell_window()
    if shell is not None:
        _focus_control(shell)
        return {"ok": True, "java_shell_detected": True, "title": _control_text(shell)}

    opened = _open_java_shell_from_gui()
    shell = _find_java_shell_window() if opened else None
    if shell is not None:
        _focus_control(shell)
        return {"ok": True, "java_shell_detected": True, "title": _control_text(shell)}

    return {
        "ok": False,
        "java_shell_detected": False,
        "manual_action": "Open COMSOL Desktop > Home > Windows > Java Shell, then retry.",
    }


@mcp.tool()
def execute_java_shell(
    code: str, allow_non_model_code: bool = False, timeout_sec: float = 30
) -> dict[str, Any]:
    """Paste Java API commands into COMSOL Java Shell and execute them with Ctrl+Enter."""
    _ensure_runtime_ready()
    executable_lines = _validate_java_shell_code(code, allow_non_model_code)
    result = _execute_in_shell(code, timeout_sec)
    result.update(
        {
            "ok": True,
            "line_count": len(code.splitlines()),
            "executable_line_count": len(executable_lines),
            "allow_non_model_code": allow_non_model_code,
        }
    )
    return result


@mcp.tool()
def set_global_parameter(name: str, value: str, description: str | None = None) -> dict[str, Any]:
    """Set a global COMSOL parameter in the current GUI model through Java Shell."""
    if not re.match(r"^[A-Za-z_]\w*$", name or ""):
        raise ValueError("Parameter name must be a valid COMSOL-style identifier, e.g. codex_gui_mcp_probe.")
    lines = [f"model.param().set({_java_string(name)}, {_java_string(value)});"]
    if description:
        lines.append(f"model.param().descr({_java_string(name)}, {_java_string(description)});")
    code = "\n".join(lines)
    result = execute_java_shell(code=code, allow_non_model_code=False, timeout_sec=10)
    result.update({"parameter": name, "value": value, "description": description, "code": code})
    return result


@mcp.tool()
def get_java_shell_output() -> dict[str, Any]:
    """Best-effort read of visible Java Shell text."""
    _ensure_runtime_ready()
    shell = _find_java_shell_window()
    if shell is None:
        return {"ok": False, "output_available": False, "message": "Java Shell is not visible."}

    texts: list[str] = []
    for control in _find_edit_like_descendants(shell):
        text = _control_text(control)
        if text and text not in texts:
            texts.append(text)
    return {
        "ok": True,
        "output_available": bool(texts),
        "texts": texts,
        "message": None if texts else "COMSOL Java Shell output may not be exposed through Windows UI Automation.",
    }


@mcp.tool()
def capture_comsol_window() -> dict[str, Any]:
    """Capture the current COMSOL main window to a PNG file."""
    top = _connect_top_window()
    result = _capture_control(top, "comsol_window")
    result["target"] = "main_window"
    result["title"] = _control_text(top)
    return result


@mcp.tool()
def capture_graphics_panel() -> dict[str, Any]:
    """Capture the visible COMSOL graphics panel to a PNG file."""
    graphics = _find_graphics_control()
    if graphics is None:
        raise RuntimeError("Could not locate the COMSOL graphics panel. Use capture_comsol_window() instead.")
    result = _capture_control(graphics, "comsol_graphics")
    result["target"] = "graphics_panel"
    result["title"] = _control_text(graphics)
    return result


def main() -> None:
    _ensure_runtime_ready()
    mcp.run()


if __name__ == "__main__":
    main()
