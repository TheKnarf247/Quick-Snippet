import sys
import ctypes
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon

from app.theme_manager import app_icon_path, set_windows_title_bar_theme

from app.constants import ensure_app_files

ensure_app_files()
from app.config_manager import ConfigManager
from app.constants import CONFIG_FILE


def ensure_default_data() -> None:
    config_manager = ConfigManager(CONFIG_FILE)
    config = config_manager.load()

    if not config.sections():
        config.add_section("Category 1")
        config.set("Category 1", "item1_title", "Example Snippet")
        config.set("Category 1", "item1_content", "Example text goes here.")
        config_manager.save()
        
def set_app_user_model_id():
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "Knarf.QuickSnippet.1"
        )
    except Exception:
        pass


def create_application():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    icon_file = app_icon_path()
    if icon_file:
        app.setWindowIcon(QIcon(icon_file))

    return app


def run_app(window_class, add_menu_bar_fn, ensure_default_data_fn):
    ensure_default_data_fn()
    set_app_user_model_id()

    app = create_application()

    window = window_class()
    add_menu_bar_fn(window)

    current_theme = window.settings.value("App/theme", "dark", type=str)
    set_windows_title_bar_theme(window, current_theme == "dark")

    sys.exit(app.exec_())