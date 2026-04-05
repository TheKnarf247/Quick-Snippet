import os
from pathlib import Path


APP_NAME = "Quick Snippet"
APP_DIR = Path(os.getenv("APPDATA") or Path.home()) / APP_NAME
CONFIG_FILE = APP_DIR / "QuickSnippetConfig.ini"
SETTINGS_FILE = APP_DIR / "settings.ini"

def ensure_app_files():
    APP_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.touch(exist_ok=True)
    SETTINGS_FILE.touch(exist_ok=True)