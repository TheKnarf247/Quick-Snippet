import os
import sys
from PyQt5.QtWidgets import (
    QMainWindow,
    QPushButton,
    QMenu,
    QApplication,
    QAction,
    QSplitter,
    QListWidget,
    QMessageBox,
    QInputDialog,
    QPlainTextEdit,
    QListWidgetItem,
    QLineEdit,
    QComboBox,
    QToolButton,
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
)
from _internal.version import __version__
from PyQt5.QtCore import QSettings, QEvent, Qt, QSize
from PyQt5.QtGui import QIcon
from pathlib import Path

try:
    from win32com.client import Dispatch
except Exception:
    Dispatch = None

from app.theme_manager import (
    icon_path,
    app_icon_path,
)

from app.widgets import *
from app.config_manager import ConfigManager
from app.theme_manager import set_windows_title_bar_theme
from app.tray_manager import TrayManager
from app.import_export import ImportExportManager
from app.constants import SETTINGS_FILE, CONFIG_FILE, APP_NAME


class QuickSnippetWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings(str(SETTINGS_FILE), QSettings.IniFormat)
        self.config_manager = ConfigManager(CONFIG_FILE)
        self.config = self.config_manager.load()
        self.current_category = None
        self.current_note_title = None
        self.is_editing = False
        self.is_quitting = False
        self.current_theme = self.settings.value("App/theme", "dark", type=str)
        self.tray_manager = TrayManager(self, APP_NAME)
        self.tray_icon = None
        self.import_export_manager = ImportExportManager(self, APP_NAME)

        self.category_sort_mode = "last_used"
        self.category_sort_ascending = True
        self.snippet_sort_mode = "az"
        self.snippet_sort_ascending = True
        self.snippet_filter_text = ""

        self.setup_defaults()
        self.build_ui()
        self.apply_theme(self.current_theme)
        self.tray_icon = self.tray_manager.create_tray_icon()
        self.load_categories()
        self.apply_startup_state()

    def refresh_button_icons(self):
        for button in self.findChildren(QPushButton):
            icon_filename = button.property("icon_filename")
            if icon_filename:
                button.setIcon(QIcon(icon_path(icon_filename, self.current_theme)))

        for button in self.findChildren(QToolButton):
            icon_filename = button.property("icon_filename")
            if icon_filename:
                button.setIcon(QIcon(icon_path(icon_filename, self.current_theme)))

    def force_exit(self):
        if hasattr(self, "tray_icon") and self.tray_icon:
            self.tray_icon.hide()
        QApplication.quit()

    def populate_settings_menu(self, menu: QMenu):
        menu.clear()

        start_with_windows_action = QAction("Start with Windows", self)
        start_with_windows_action.setCheckable(True)
        start_with_windows_action.setChecked(
            self.settings.value("App/start_with_windows", False, type=bool)
        )
        start_with_windows_action.triggered.connect(self.toggle_start_with_windows)

        start_minimized_action = QAction("Start minimized", self)
        start_minimized_action.setCheckable(True)
        start_minimized_action.setChecked(
            self.settings.value("App/start_minimized", False, type=bool)
        )
        start_minimized_action.triggered.connect(
            lambda checked: self.save_setting("App/start_minimized", checked)
        )

        show_copy_notification_action = QAction("Show text copied notification", self)
        show_copy_notification_action.setCheckable(True)
        show_copy_notification_action.setChecked(
            self.settings.value("App/show_copy_notification", True, type=bool)
        )
        show_copy_notification_action.triggered.connect(
            lambda checked: self.save_setting("App/show_copy_notification", checked)
        )

        minimize_on_close_action = QAction(
            "Minimize to tray when window is closed", self
        )
        minimize_on_close_action.setCheckable(True)
        minimize_on_close_action.setChecked(
            self.settings.value("App/minimize_to_tray_on_close", True, type=bool)
        )
        minimize_on_close_action.triggered.connect(
            lambda checked: self.save_setting("App/minimize_to_tray_on_close", checked)
        )

        minimize_on_minimize_action = QAction(
            "Minimize to tray when window is minimized", self
        )
        minimize_on_minimize_action.setCheckable(True)
        minimize_on_minimize_action.setChecked(
            self.settings.value("App/minimize_to_tray_on_minimize", False, type=bool)
        )
        minimize_on_minimize_action.triggered.connect(
            lambda checked: self.save_setting("App/minimize_to_tray_on_minimize", checked)
        )

        menu.addAction(start_with_windows_action)
        menu.addAction(start_minimized_action)
        menu.addAction(show_copy_notification_action)
        menu.addSeparator()
        menu.addAction(minimize_on_close_action)
        menu.addAction(minimize_on_minimize_action)

    def get_category_snippets(self, category: str) -> list[dict]:
        snippets = []
        if category not in self.config:
            return snippets

        for key, value in self.config[category].items():
            if key.endswith("_title"):
                content_key = key.replace("_title", "_content")
                snippets.append({
                    "title": value,
                    "content": self.config[category].get(content_key, "")
                })

        return snippets

    def export_all_data(self):
        self.import_export_manager.export_all_data()

    def import_full_backup(self, data: dict):
        self.import_export_manager.import_full_backup(data)

    def export_category(self):
        self.import_export_manager.export_category()

    def export_note(self):
        self.import_export_manager.export_note()

    def import_data(self):
        self.import_export_manager.import_data()

    def make_unique_category_name(self, base_name: str) -> str:
        if base_name not in self.config.sections():
            return base_name

        counter = 1
        while True:
            candidate = f"{base_name} ({counter})"
            if candidate not in self.config.sections():
                return candidate
            counter += 1

    def make_unique_note_title(self, category: str, base_title: str) -> str:
        if not self.note_exists(category, base_title):
            return base_title

        counter = 1
        while True:
            candidate = f"{base_title} ({counter})"
            if not self.note_exists(category, candidate):
                return candidate
            counter += 1

    def setup_defaults(self):
        defaults = {
            "App/start_with_windows": False,
            "App/start_minimized": False,
            "App/show_copy_notification": True,
            "App/minimize_to_tray_on_close": True,
            "App/minimize_to_tray_on_minimize": False,
            "App/theme": "dark",
        }
        for key, value in defaults.items():
            if self.settings.value(key) is None:
                self.settings.setValue(key, value)
        self.settings.sync()

    def build_ui(self):
        self.setWindowTitle(APP_NAME)

        icon_file = app_icon_path()
        if icon_file:
            self.setWindowIcon(QIcon(icon_file))

        self.resize(1200, 720)
        self.setMinimumSize(900, 560)

        central = QWidget()
        root = QHBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.build_category_panel())
        splitter.addWidget(self.build_note_panel())
        splitter.addWidget(self.build_content_panel())
        splitter.setSizes([220, 360, 520])
        root.addWidget(splitter)
        self.setCentralWidget(central)

        self.statusBar().showMessage("Ready")

    def build_panel_header(self, title: str, buttons):
        wrapper = QWidget()
        layout = QVBoxLayout(wrapper)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        toolbar = QWidget()
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(8, 6, 8, 2)
        toolbar_layout.setSpacing(6)

        for button in buttons:
            toolbar_layout.addWidget(button)
        toolbar_layout.addStretch(1)

        title_label = QLabel(title)
        title_label.setStyleSheet(
            """
            font-size: 16px;
            font-weight: 600;
            padding: 2px 8px 4px 8px;
        """
        )

        layout.addWidget(toolbar)
        layout.addWidget(title_label)
        return wrapper, layout, title_label

    def make_small_button(self, icon_filename: str, slot, tooltip: str = ""):
        button = QPushButton()
        button.setCursor(Qt.PointingHandCursor)
        button.setFixedSize(48, 48)
        button.setProperty("icon_filename", icon_filename)
        button.setIcon(QIcon(icon_path(icon_filename, self.current_theme)))
        button.setIconSize(QSize(28, 28))
        button.clicked.connect(slot)

        if tooltip:
            button.setToolTip(tooltip)

        return button

    def build_category_panel(self):
        add_btn = self.make_small_button(
            "IconAdd.png", self.add_category, "Add category"
        )
        delete_btn = self.make_small_button(
            "IconRemove.png", self.delete_category, "Delete category"
        )
        rename_btn = self.make_small_button(
            "IconRename.png", self.rename_category, "Rename category"
        )
        wrapper, layout, self.category_header_label = self.build_panel_header(
            "Categories", [add_btn, delete_btn, rename_btn]
        )

        self.category_list = QListWidget()

        font = self.category_list.font()
        font.setPointSize(14)
        font.setBold(True)
        self.category_list.setFont(font)

        self.category_list.itemSelectionChanged.connect(self.on_category_changed)
        layout.addWidget(self.category_list, 1)
        return wrapper

    def build_note_panel(self):
        add_btn = self.make_small_button("IconAdd.png", self.add_note, "Add snippet")
        delete_btn = self.make_small_button(
            "IconRemove.png", self.delete_note, "Delete snippet"
        )
        rename_btn = self.make_small_button(
            "IconRename.png", self.rename_note, "Rename snippet"
        )
        move_btn = self.make_small_button(
            "IconMove.png", self.move_note, "Move snippet"
        )
        edit_btn = self.make_small_button(
            "IconEdit.png", self.start_editing_content, "Edit snippet"
        )
        wrapper, layout, self.note_header_label = self.build_panel_header(
            "Snippets", [add_btn, delete_btn, rename_btn, move_btn, edit_btn]
        )

        controls_row = QWidget()
        controls_layout = QHBoxLayout(controls_row)
        controls_layout.setContentsMargins(8, 0, 8, 4)
        controls_layout.setSpacing(6)

        self.snippet_filter_input = QLineEdit()
        self.snippet_filter_input.setPlaceholderText("Filter snippets...")
        self.snippet_filter_input.textChanged.connect(self.on_snippet_filter_changed)

        self.snippet_clear_filter_btn = QToolButton()
        self.snippet_clear_filter_btn.setText("✕")
        self.snippet_clear_filter_btn.setToolTip("Clear filter")
        self.snippet_clear_filter_btn.clicked.connect(self.clear_snippet_filter)

        self.snippet_sort_combo = QComboBox()
        self.snippet_sort_combo.addItem("A-Z", "az")
        self.snippet_sort_combo.currentIndexChanged.connect(self.on_snippet_sort_changed)

        self.snippet_sort_direction_btn = QToolButton()
        self.snippet_sort_direction_btn.setText("↑")
        self.snippet_sort_direction_btn.setToolTip("Toggle ascending/descending")
        self.snippet_sort_direction_btn.clicked.connect(self.toggle_snippet_sort_direction)

        controls_layout.addWidget(self.snippet_filter_input, 1)
        controls_layout.addWidget(self.snippet_clear_filter_btn)
        controls_layout.addWidget(self.snippet_sort_combo)
        controls_layout.addWidget(self.snippet_sort_direction_btn)

        self.note_list = QListWidget()
        self.note_list.itemSelectionChanged.connect(self.on_note_changed)

        layout.addWidget(controls_row)
        layout.addWidget(self.note_list, 1)
        return wrapper

    def on_snippet_filter_changed(self, text: str):
        self.snippet_filter_text = text.strip()
        self.load_notes()

    def clear_snippet_filter(self):
        self.snippet_filter_input.clear()

    def on_snippet_sort_changed(self):
        self.snippet_sort_mode = self.snippet_sort_combo.currentData()
        self.load_notes()

    def toggle_snippet_sort_direction(self):
        self.snippet_sort_ascending = not self.snippet_sort_ascending
        self.snippet_sort_direction_btn.setText("↑" if self.snippet_sort_ascending else "↓")
        self.load_notes()

    def get_sorted_filtered_note_titles(self, category: str) -> list[str]:
        note_titles = []

        if not category or category not in self.config:
            return note_titles

        for key, value in self.config[category].items():
            if key.endswith("_title"):
                note_titles.append(value)

        filter_text = self.snippet_filter_text.lower()
        if filter_text:
            note_titles = [title for title in note_titles if filter_text in title.lower()]

        if self.snippet_sort_mode == "az":
            note_titles.sort(key=lambda title: title.lower(), reverse=not self.snippet_sort_ascending)

        return note_titles

    def move_note(self):
        if not self.current_category or not self.current_note_title:
            QMessageBox.information(self, APP_NAME, "Please select a snippet to move.")
            return

        available_categories = [
            category
            for category in self.config.sections()
            if category != self.current_category
        ]

        if not available_categories:
            QMessageBox.information(
                self, APP_NAME, "Create another category before moving a snippet."
            )
            return

        target_category, ok = QInputDialog.getItem(
            self,
            "Move Snippet",
            f"Move '{self.current_note_title}' to:",
            available_categories,
            0,
            False,
        )

        if not ok or not target_category:
            return

        if self.note_exists(target_category, self.current_note_title):
            QMessageBox.warning(
                self,
                APP_NAME,
                f"A snippet named '{self.current_note_title}' already exists in '{target_category}'.",
            )
            return

        source_title_key = None
        source_content_key = None
        content = ""

        for key, value in self.config[self.current_category].items():
            if key.endswith("_title") and value == self.current_note_title:
                source_title_key = key
                source_content_key = key.replace("_title", "_content")
                content = self.config[self.current_category].get(source_content_key, "")
                break

        if source_title_key is None or source_content_key is None:
            QMessageBox.warning(self, APP_NAME, "Could not find the selected snippet.")
            return

        new_index = self.get_next_item_index(target_category)
        self.config[target_category][f"item{new_index}_title"] = self.current_note_title
        self.config[target_category][f"item{new_index}_content"] = content

        self.config.remove_option(self.current_category, source_title_key)
        self.config.remove_option(self.current_category, source_content_key)

        self.reindex_category(self.current_category)
        self.reindex_category(target_category)
        self.config_manager.save()

        moved_title = self.current_note_title
        self.current_note_title = None

        self.load_categories()
        self.select_category_by_name(target_category)
        self.load_notes()
        self.select_note_by_title(moved_title)

    def build_content_panel(self):
        edit_btn = self.make_small_button(
            "IconEdit.png", self.start_editing_content, "Edit content"
        )
        save_btn = self.make_small_button(
            "IconSave.png", self.finish_editing_content, "Save content"
        )
        wrapper, layout, self.content_header_label = self.build_panel_header(
            "Select a snippet", [edit_btn, save_btn]
        )

        self.content_editor = QPlainTextEdit()
        self.content_editor.setReadOnly(True)
        self.content_editor.setPlaceholderText("Select a snippet to view its content.")

        layout.addWidget(self.content_editor, 1)

        self.content_edit_button = edit_btn
        self.content_save_button = save_btn
        self.content_save_button.setEnabled(False)
        return wrapper

    def set_theme(self, theme_name: str):
        self.current_theme = theme_name
        self.save_setting("App/theme", theme_name)
        self.apply_theme(theme_name)
        self.refresh_button_icons()
        self.load_notes()

    def apply_theme(self, theme_name: str):
        self.current_theme = theme_name
        if theme_name == "light":
            self.setStyleSheet(
                """
                QMainWindow, QWidget {
                    background-color: #f3f4f7;
                    color: #2b2f36;
                }

                QLabel {
                    color: #2b2f36;
                }

                QListWidget, QPlainTextEdit, QLineEdit, QComboBox {
                    background-color: #ffffff;
                    color: #4a4f57;
                    border: 1px solid #cfd4dc;
                    border-radius: 8px;
                    padding: 4px;
                    font-size: 14px;
                }

                QLabel#snippetPreviewLabel {
                    background: transparent;
                    color: #6b7280;
                    font-size: 11px;
                    font-weight: 400;
                    padding: 0px;
                    margin: 0px;
                    border: none;
                }

                QListWidget::item {
                    padding: 0px;
                    border-radius: 8px;
                }

                QListWidget::item:selected {
                    background-color: #cfe3ff;
                    color: #1f2937;
                }

                QLabel#snippetTitleLabel {
                    background: transparent;
                    color: #4a4f57;
                    font-size: 15px;
                    font-weight: 800;
                    padding: 0px;
                    margin: 0px;
                    border: none;
                }

                QPushButton, QToolButton {
                    background-color: #ffffff;
                    color: #2b2f36;
                    border: 1px solid #cfd4dc;
                    border-radius: 8px;
                    padding: 2px;
                    font-size: 12px;
                }

                QListWidget::item {
                    font-weight: 800;
                }

                NoteListItemWidget {
                    background: transparent;
                    border: none;
                }

                QPushButton:hover, QToolButton:hover {
                    background-color: #eef4ff;
                    border: 1px solid #aac7f2;
                }

                QPushButton:pressed, QToolButton:pressed {
                    background-color: #dde9f8;
                }

                QMenuBar {
                    background-color: #e9edf3;
                    color: #2b2f36;
                }

                QMenuBar::item:selected {
                    background-color: #cfe3ff;
                    border-radius: 4px;
                }

                QMenu {
                    background-color: #ffffff;
                    color: #2b2f36;
                    border: 1px solid #cfd4dc;
                }

                QMenu::item:selected {
                    background-color: #cfe3ff;
                }

                QStatusBar {
                    background-color: #e9edf3;
                    color: #2b2f36;
                }
            """
            )
            set_windows_title_bar_theme(self, False)
        else:
            self.setStyleSheet(
                """
                QMainWindow, QWidget {
                    background-color: #1e1f22;
                    color: #e8eaed;
                }

                QLabel {
                    color: #e8eaed;
                }

                QListWidget, QPlainTextEdit, QLineEdit, QComboBox {
                    background-color: #2b2d31;
                    color: #d6d9df;
                    border: 1px solid #3c4043;
                    border-radius: 8px;
                    padding: 4px;
                    font-size: 14px;
                }

                QLabel#snippetPreviewLabel {
                    background: transparent;
                    color: #aeb6c2;
                    font-size: 11px;
                    font-weight: 400;
                    padding: 0px;
                    margin: 0px;
                    border: none;
                }

                QListWidget::item {
                    padding: 0px;
                    border-radius: 8px;
                }

                QListWidget::item:selected {
                    background-color: #375a7f;
                    color: #e6e9ee;
                }

                QLabel#snippetTitleLabel {
                    background: transparent;
                    color: #d6d9df;
                    font-size: 15px;
                    font-weight: 800;
                    padding: 0px;
                    margin: 0px;
                    border: none;
                }

                QPushButton, QToolButton {
                    background-color: #2b2d31;
                    color: #e8eaed;
                    border: 1px solid #3c4043;
                    border-radius: 8px;
                    padding: 2px;
                    font-size: 12px;
                }

                QListWidget::item {
                    font-weight: 800;
                }

                NoteListItemWidget {
                    background: transparent;
                    border: none;
                }

                QPushButton:hover, QToolButton:hover {
                    background-color: #444;
                    border: 1px solid #5a5a5a;
                }

                QPushButton:pressed, QToolButton:pressed {
                    background-color: #2a2a2a;
                }

                QMenuBar {
                    background-color: #232428;
                    color: #e8eaed;
                }

                QMenuBar::item:selected {
                    background-color: #375a7f;
                    border-radius: 4px;
                }

                QMenu {
                    background-color: #2b2d31;
                    color: #e8eaed;
                    border: 1px solid #3c4043;
                }

                QMenu::item:selected {
                    background-color: #375a7f;
                }

                QStatusBar {
                    background-color: #232428;
                    color: #e8eaed;
                }
            """
            )
            set_windows_title_bar_theme(self, True)

    def apply_startup_state(self):
        if (
            self.settings.value("App/start_minimized", False, type=bool)
            or "--minimized" in sys.argv
        ):
            self.hide()
        else:
            self.show()

    def changeEvent(self, event):
        if event.type() == QEvent.WindowStateChange:
            if self.isMinimized() and self.settings.value(
                "App/minimize_to_tray_on_minimize", False, type=bool
            ):
                event.ignore()
                self.hide()
                return
        super().changeEvent(event)

    def closeEvent(self, event):
        if self.is_quitting:
            event.accept()
            return

        if self.settings.value("App/minimize_to_tray_on_close", True, type=bool):
            event.ignore()
            self.hide()

            if self.tray_icon:
                self.tray_icon.showMessage(
                    APP_NAME,
                    "App minimized to tray",
                    self.tray_icon.Information,
                    2000,
                )
            return

        event.accept()

    def restore_from_tray(self):
        self.tray_manager.restore_from_tray()

    def exit_application(self):
        self.tray_manager.exit_application()

    def notify(self, message: str):
        self.tray_manager.notify(message)

    def load_categories(self):
        self.config = self.config_manager.load()
        self.category_list.clear()
        for section in self.config.sections():
            self.category_list.addItem(section)

        if self.category_list.count() > 0:
            self.category_list.setCurrentRow(0)
        else:
            self.current_category = None
            self.load_notes()

    def load_notes(self):
        selected_title = self.current_note_title

        self.note_list.clear()
        self.content_header_label.setText("Select a snippet")
        self.content_editor.clear()
        self.content_editor.setReadOnly(True)
        self.current_note_title = None
        self.is_editing = False
        self.content_save_button.setEnabled(False)

        if not self.current_category or self.current_category not in self.config:
            self.note_header_label.setText("Snippets")
            self.category_header_label.setText("Categories")
            return

        self.category_header_label.setText("Categories")
        self.note_header_label.setText(f"{self.current_category} Snippets")

        note_titles = self.get_sorted_filtered_note_titles(self.current_category)

        for title in note_titles:
            item = QListWidgetItem()
            item.setData(Qt.UserRole, title)
            item.setSizeHint(QSize(100, 76))
            self.note_list.addItem(item)

            preview = self.get_note_content(self.current_category, title)

            widget = NoteListItemWidget(
                title,
                preview,
                lambda checked=False, t=title: self.copy_note_by_title(t),
                icon_path,
                self.current_theme,
            )

            self.note_list.setItemWidget(item, widget)

        if selected_title:
            self.select_note_by_title(selected_title)

        if self.note_list.count() > 0 and self.note_list.currentRow() < 0:
            self.note_list.setCurrentRow(0)

    def on_category_changed(self):
        items = self.category_list.selectedItems()
        self.current_category = items[0].text() if items else None
        self.load_notes()

    def on_note_changed(self):
        item = self.note_list.currentItem()
        if not item or not self.current_category:
            self.current_note_title = None
            self.content_header_label.setText("Select a snippet")
            self.content_editor.clear()
            return

        self.current_note_title = item.data(Qt.UserRole)
        self.content_header_label.setText(self.current_note_title)
        self.content_editor.setPlainText(
            self.get_note_content(self.current_category, self.current_note_title)
        )
        self.content_editor.setReadOnly(True)
        self.is_editing = False
        self.content_save_button.setEnabled(False)

    def get_next_item_index(self, category: str) -> int:
        existing = []
        if category in self.config:
            for key in self.config[category].keys():
                if key.startswith("item") and key.endswith("_title"):
                    try:
                        existing.append(int(key[4:-6]))
                    except ValueError:
                        pass
        return max(existing, default=0) + 1

    def get_note_content(self, category: str, title: str) -> str:
        if category not in self.config:
            return ""
        for key, value in self.config[category].items():
            if key.endswith("_title") and value == title:
                return self.config[category].get(key.replace("_title", "_content"), "")
        return ""

    def set_note_content(self, category: str, title: str, content: str) -> None:
        for key, value in self.config[category].items():
            if key.endswith("_title") and value == title:
                self.config[category][key.replace("_title", "_content")] = content
                self.config_manager.save()
                return

    def add_category(self):
        name, ok = QInputDialog.getText(self, APP_NAME, "New category name:")
        if not ok or not name.strip():
            return
        name = name.strip()
        if name in self.config.sections():
            QMessageBox.warning(self, APP_NAME, "That category already exists.")
            return

        self.config.add_section(name)
        self.config_manager.save()
        self.load_categories()
        matches = self.category_list.findItems(name, Qt.MatchExactly)
        if matches:
            self.category_list.setCurrentItem(matches[0])

    def rename_category(self):
        if not self.current_category:
            return
        new_name, ok = QInputDialog.getText(
            self, APP_NAME, "Rename category:", text=self.current_category
        )
        if not ok or not new_name.strip():
            return
        new_name = new_name.strip()
        if new_name == self.current_category:
            return
        if new_name in self.config.sections():
            QMessageBox.warning(self, APP_NAME, "That category already exists.")
            return

        self.config.add_section(new_name)
        for key, value in self.config[self.current_category].items():
            self.config[new_name][key] = value
        self.config.remove_section(self.current_category)
        self.config_manager.save()
        self.load_categories()
        matches = self.category_list.findItems(new_name, Qt.MatchExactly)
        if matches:
            self.category_list.setCurrentItem(matches[0])

    def delete_category(self):
        if not self.current_category:
            return
        if (
            QMessageBox.question(
                self, APP_NAME, f"Delete category '{self.current_category}'?"
            )
            != QMessageBox.Yes
        ):
            return
        self.config.remove_section(self.current_category)
        if not self.config.sections():
            self.config.add_section("Category 1")
        self.config_manager.save()
        self.load_categories()

    def add_note(self):
        if not self.current_category:
            QMessageBox.warning(self, APP_NAME, "Create or select a category first.")
            return
        title, ok = QInputDialog.getText(self, APP_NAME, "New snippet name:")
        if not ok or not title.strip():
            return
        title = title.strip()
        if self.note_exists(self.current_category, title):
            QMessageBox.warning(
                self, APP_NAME, "That snippet already exists in this category."
            )
            return

        index = self.get_next_item_index(self.current_category)
        self.config[self.current_category][f"item{index}_title"] = title
        self.config[self.current_category][f"item{index}_content"] = ""
        self.config_manager.save()
        self.load_notes()
        self.select_note_by_title(title)
        self.start_editing_content()

    def rename_note(self):
        if not self.current_category or not self.current_note_title:
            return
        new_title, ok = QInputDialog.getText(
            self, APP_NAME, "Rename snippet:", text=self.current_note_title
        )
        if not ok or not new_title.strip():
            return
        new_title = new_title.strip()
        if new_title == self.current_note_title:
            return
        if self.note_exists(self.current_category, new_title):
            QMessageBox.warning(
                self, APP_NAME, "That snippet already exists in this category."
            )
            return

        for key, value in self.config[self.current_category].items():
            if key.endswith("_title") and value == self.current_note_title:
                self.config[self.current_category][key] = new_title
                break
        self.config_manager.save()
        self.load_notes()
        self.select_note_by_title(new_title)

    def delete_note(self):
        if not self.current_category or not self.current_note_title:
            return
        if (
            QMessageBox.question(
                self, APP_NAME, f"Delete snippet '{self.current_note_title}'?"
            )
            != QMessageBox.Yes
        ):
            return

        keys_to_remove = []
        for key, value in self.config[self.current_category].items():
            if key.endswith("_title") and value == self.current_note_title:
                keys_to_remove.append(key)
                keys_to_remove.append(key.replace("_title", "_content"))
                break

        for key in keys_to_remove:
            self.config.remove_option(self.current_category, key)

        self.reindex_category(self.current_category)
        self.config_manager.save()
        self.load_notes()

    def reindex_category(self, category: str):
        pairs = []
        for key, value in list(self.config[category].items()):
            if key.endswith("_title"):
                content_key = key.replace("_title", "_content")
                pairs.append((value, self.config[category].get(content_key, "")))

        for key in list(self.config[category].keys()):
            self.config.remove_option(category, key)

        for idx, (title, content) in enumerate(pairs, start=1):
            self.config[category][f"item{idx}_title"] = title
            self.config[category][f"item{idx}_content"] = content

    def note_exists(self, category: str, title: str) -> bool:
        for key, value in self.config[category].items():
            if key.endswith("_title") and value == title:
                return True
        return False

    def select_category_by_name(self, category_name):
        for i in range(self.category_list.count()):
            item = self.category_list.item(i)
            if item and item.text() == category_name:
                self.category_list.setCurrentRow(i)
                return

    def select_note_by_title(self, note_title):
        for i in range(self.note_list.count()):
            item = self.note_list.item(i)
            if item and item.data(Qt.UserRole) == note_title:
                self.note_list.setCurrentRow(i)
                return

    def start_editing_content(self):
        if not self.current_note_title:
            return
        self.is_editing = True
        self.content_editor.setReadOnly(False)
        self.content_editor.setFocus()
        self.content_save_button.setEnabled(True)
        self.statusBar().showMessage(f"Editing: {self.current_note_title}")

    def finish_editing_content(self):
        if not (self.current_category and self.current_note_title and self.is_editing):
            return

        saved_title = self.current_note_title

        self.set_note_content(
            self.current_category,
            self.current_note_title,
            self.content_editor.toPlainText(),
        )

        self.is_editing = False
        self.content_editor.setReadOnly(True)
        self.content_save_button.setEnabled(False)

        self.load_notes()
        self.select_note_by_title(saved_title)

        self.statusBar().showMessage("Snippet saved.", 3000)

    def copy_note_by_title(self, title: str):
        if not self.current_category:
            return
        content = self.get_note_content(self.current_category, title)
        QApplication.clipboard().setText(content)
        self.notify("Text copied to clipboard")

    def save_setting(self, key: str, value):
        self.settings.setValue(key, value)
        self.settings.sync()

    def toggle_start_with_windows(self, checked: bool):
        startup_folder = (
            Path(os.getenv("APPDATA") or "")
            / r"Microsoft\Windows\Start Menu\Programs\Startup"
        )
        shortcut_path = startup_folder / f"{APP_NAME}.lnk"
        exe_path = os.path.abspath(sys.argv[0])

        if checked:
            if Dispatch is None:
                QMessageBox.warning(
                    self,
                    APP_NAME,
                    "Startup feature unavailable (win32com not installed).",
                )
                return

            shell = Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(str(shortcut_path))
            shortcut.TargetPath = exe_path
            shortcut.WorkingDirectory = os.path.dirname(exe_path)
            shortcut.Arguments = "--minimized"
            shortcut.Save()
        else:
            if shortcut_path.exists():
                shortcut_path.unlink()

        self.save_setting("App/start_with_windows", checked)


def add_menu_bar(window: "QuickSnippetWindow"):
    menu_bar = window.menuBar()

    file_menu = menu_bar.addMenu("File")

    import_action = QAction("Import from JSON...", window)
    import_action.triggered.connect(window.import_data)
    file_menu.addAction(import_action)

    file_menu.addSeparator()

    export_all_action = QAction("Export all categories...", window)
    export_all_action.triggered.connect(window.export_all_data)
    file_menu.addAction(export_all_action)

    export_category_action = QAction("Export current category...", window)
    export_category_action.triggered.connect(window.export_category)
    file_menu.addAction(export_category_action)

    export_snippet_action = QAction("Export selected snippet...", window)
    export_snippet_action.triggered.connect(window.export_note)
    file_menu.addAction(export_snippet_action)

    file_menu.addSeparator()

    settings_menu = file_menu.addMenu("Settings")
    window.populate_settings_menu(settings_menu)

    file_menu.addSeparator()

    exit_action = QAction("Exit", window)
    exit_action.setShortcut("Ctrl+Q")
    exit_action.triggered.connect(window.force_exit)
    file_menu.addAction(exit_action)

    appearance_menu = menu_bar.addMenu("Appearance")

    light_action = QAction("Light", window)
    light_action.triggered.connect(lambda: window.set_theme("light"))
    appearance_menu.addAction(light_action)

    dark_action = QAction("Dark", window)
    dark_action.triggered.connect(lambda: window.set_theme("dark"))
    appearance_menu.addAction(dark_action)

    about_menu = menu_bar.addMenu("About")

    about_action = QAction(f"About {APP_NAME}", window)
    about_action.triggered.connect(
        lambda: QMessageBox.information(
            window,
            f"About {APP_NAME} 2026",
            f"{APP_NAME} v{__version__}\n2026\n\nReusable text snippets for Windows.\n\nBy Frank Noonan",
        )
    )
    about_menu.addAction(about_action)