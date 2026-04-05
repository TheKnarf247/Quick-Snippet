import sys
import ctypes
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon

from app.theme_manager import app_icon_path, set_windows_title_bar_theme


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