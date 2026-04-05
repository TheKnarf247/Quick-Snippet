import os
import sys

from PyQt5.QtWidgets import QAction, QApplication, QMenu, QStyle, QSystemTrayIcon
from PyQt5.QtGui import QIcon

from app.theme_manager import app_icon_path


class TrayManager:
    def __init__(self, window, app_name: str):
        self.window = window
        self.app_name = app_name
        self.tray_icon = None

    def create_tray_icon(self):
        icon_file = app_icon_path()
        icon = (
            QIcon(icon_file)
            if icon_file
            else self.window.style().standardIcon(QStyle.SP_FileDialogDetailedView)
        )

        self.tray_icon = QSystemTrayIcon(icon, self.window)
        self.tray_icon.setToolTip(self.app_name)

        menu = QMenu(self.window)

        open_action = QAction(f"Open {self.app_name}", self.window)
        open_action.triggered.connect(self.restore_from_tray)
        menu.addAction(open_action)

        restart_action = QAction(f"Restart {self.app_name}", self.window)
        restart_action.triggered.connect(self.restart_app)
        menu.addAction(restart_action)

        menu.addSeparator()

        exit_action = QAction(f"Exit {self.app_name}", self.window)
        exit_action.triggered.connect(self.exit_application)
        menu.addAction(exit_action)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(self.on_tray_icon_activated)
        self.tray_icon.show()

        return self.tray_icon

    def on_tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.restore_from_tray()

    def restore_from_tray(self):
        self.window.showNormal()
        self.window.raise_()
        self.window.activateWindow()

    def restart_app(self):
        QApplication.quit()
        os.execl(sys.executable, sys.executable, *sys.argv)

    def exit_application(self):
        self.window.is_quitting = True
        if self.tray_icon:
            self.tray_icon.hide()

        app = QApplication.instance()
        if app is not None:
            app.quit()

    def notify(self, message: str):
        self.window.statusBar().showMessage(message, 3000)
        if self.window.settings.value("App/show_copy_notification", True, type=bool):
            if self.tray_icon:
                self.tray_icon.showMessage(
                    self.app_name, message, QSystemTrayIcon.Information, 2000
                )