import sys
import ctypes
from ctypes import wintypes
from pathlib import Path

def resource_path(relative_path: str) -> Path:
    base_path = getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent)
    return Path(base_path) / relative_path


def icon_path(filename: str, theme: str = "dark") -> str:
    base = resource_path("_internal/icons")

    if str(theme).lower() == "light":
        name = Path(filename).name
        if name.startswith("Icon"):
            light_name = name.replace("Icon", "IconLight", 1)
            light_path = base / light_name
            if light_path.exists():
                return str(light_path)

    return str(base / filename)


def app_icon_path() -> str:
    ico_path = resource_path("QuickSnippet.ico")
    if ico_path.exists():
        return str(ico_path)
    return ""


def set_windows_title_bar_theme(window, dark: bool):
    try:
        hwnd = int(window.winId())
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        value = ctypes.c_int(1 if dark else 0)

        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            wintypes.HWND(hwnd),
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value),
            ctypes.sizeof(value),
        )
    except Exception:
        pass

