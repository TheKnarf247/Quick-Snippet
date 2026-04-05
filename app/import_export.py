import json

from PyQt5.QtWidgets import QFileDialog, QMessageBox, QInputDialog


class ImportExportManager:
    def __init__(self, window, app_name: str):
        self.window = window
        self.app_name = app_name

    def export_all_data(self):
        categories = []

        for category in self.window.config.sections():
            categories.append({
                "category": category,
                "snippets": self.window.get_category_snippets(category),
            })

        data = {
            "type": "full_backup",
            "categories": categories,
        }

        file_path, _ = QFileDialog.getSaveFileName(
            self.window,
            "Export All Categories",
            "QuickSnippetBackup.json",
            "JSON Files (*.json)"
        )

        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2, ensure_ascii=False)
            self.window.statusBar().showMessage("All categories exported.", 3000)
        except Exception as e:
            QMessageBox.warning(self.window, self.app_name, f"Failed to export backup:\n{e}")

    def import_full_backup(self, data: dict):
        categories = data.get("categories", [])
        if not categories:
            QMessageBox.warning(self.window, self.app_name, "That backup file contains no categories.")
            return

        imported_first_category = None

        for category_data in categories:
            category_name = (category_data.get("category") or "Imported Category").strip() or "Imported Category"
            snippets = category_data.get("snippets", [])

            if category_name in self.window.config.sections():
                category_name = self.window.make_unique_category_name(category_name)

            self.window.config.add_section(category_name)

            for snippet in snippets:
                title = (snippet.get("title") or "Imported Snippet").strip() or "Imported Snippet"
                content = snippet.get("content", "")

                index = self.window.get_next_item_index(category_name)
                self.window.config[category_name][f"item{index}_title"] = title
                self.window.config[category_name][f"item{index}_content"] = content

            if imported_first_category is None:
                imported_first_category = category_name

        self.window.config_manager.save()
        self.window.load_categories()

        if imported_first_category:
            self.window.select_category_by_name(imported_first_category)
            self.window.load_notes()

        self.window.statusBar().showMessage("Backup imported.", 3000)

    def export_category(self):
        if not self.window.current_category:
            QMessageBox.information(self.window, self.app_name, "Please select a category to export.")
            return

        data = {
            "type": "category",
            "category": self.window.current_category,
            "snippets": self.window.get_category_snippets(self.window.current_category),
        }

        file_path, _ = QFileDialog.getSaveFileName(
            self.window,
            "Export Category",
            f"{self.window.current_category}.json",
            "JSON Files (*.json)"
        )

        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2, ensure_ascii=False)
            self.window.statusBar().showMessage("Category exported.", 3000)
        except Exception as e:
            QMessageBox.warning(self.window, self.app_name, f"Failed to export category:\n{e}")

    def export_note(self):
        if not self.window.current_category or not self.window.current_note_title:
            QMessageBox.information(self.window, self.app_name, "Please select a snippet to export.")
            return

        data = {
            "type": "snippet",
            "category": self.window.current_category,
            "title": self.window.current_note_title,
            "content": self.window.get_note_content(
                self.window.current_category,
                self.window.current_note_title,
            ),
        }

        file_path, _ = QFileDialog.getSaveFileName(
            self.window,
            "Export Snippet",
            f"{self.window.current_note_title}.json",
            "JSON Files (*.json)"
        )

        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2, ensure_ascii=False)
            self.window.statusBar().showMessage("Snippet exported.", 3000)
        except Exception as e:
            QMessageBox.warning(self.window, self.app_name, f"Failed to export snippet:\n{e}")

    def import_data(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self.window,
            "Import Data",
            "",
            "JSON Files (*.json)"
        )

        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception as e:
            QMessageBox.warning(self.window, self.app_name, f"Failed to read import file:\n{e}")
            return

        data_type = data.get("type")

        if data_type == "full_backup":
            self.import_full_backup(data)
            return

        if data_type == "category":
            category_name = data.get("category", "Imported Category").strip() or "Imported Category"
            snippets = data.get("snippets", [])

            if category_name in self.window.config.sections():
                msg = QMessageBox(self.window)
                msg.setWindowTitle("Import Category")
                msg.setText(f"The category '{category_name}' already exists.")
                msg.setInformativeText("What would you like to do?")
                merge_button = msg.addButton("Merge", QMessageBox.AcceptRole)
                rename_button = msg.addButton("Rename Imported Category", QMessageBox.ActionRole)
                cancel_button = msg.addButton("Cancel", QMessageBox.RejectRole)
                msg.setIcon(QMessageBox.Question)
                msg.exec_()

                clicked = msg.clickedButton()

                if clicked == cancel_button:
                    return

                if clicked == rename_button:
                    new_category, ok = QInputDialog.getText(
                        self.window,
                        "Rename Imported Category",
                        "Enter a new category name:",
                        text=f"{category_name} Imported"
                    )
                    if not ok or not new_category.strip():
                        return

                    category_name = new_category.strip()
                    if category_name in self.window.config.sections():
                        category_name = self.window.make_unique_category_name(category_name)

                    self.window.config.add_section(category_name)

                elif clicked == merge_button:
                    pass
            else:
                self.window.config.add_section(category_name)

            imported_titles = []

            for snippet in snippets:
                title = (snippet.get("title") or "Imported Snippet").strip() or "Imported Snippet"
                content = snippet.get("content", "")

                title = self.window.make_unique_note_title(category_name, title)

                index = self.window.get_next_item_index(category_name)
                self.window.config[category_name][f"item{index}_title"] = title
                self.window.config[category_name][f"item{index}_content"] = content
                imported_titles.append(title)

            self.window.config_manager.save()
            self.window.load_categories()
            self.window.select_category_by_name(category_name)
            self.window.load_notes()

            if imported_titles:
                self.window.select_note_by_title(imported_titles[0])

            self.window.statusBar().showMessage("Category imported.", 3000)
            return

        if data_type == "snippet":
            available_categories = list(self.window.config.sections())
            category_choices = available_categories + ["New Category..."]

            if not available_categories:
                new_category, ok = QInputDialog.getText(
                    self.window,
                    "Import Snippet",
                    "No categories exist yet.\nEnter a new category name:",
                    text="Category 1"
                )
                if not ok or not new_category.strip():
                    return

                target_category = new_category.strip()

                if not target_category:
                    return

                if target_category not in self.window.config.sections():
                    self.window.config.add_section(target_category)
            else:
                target_category, ok = QInputDialog.getItem(
                    self.window,
                    "Import Snippet",
                    "Choose a category:",
                    category_choices,
                    0,
                    False
                )

                if not ok or not target_category:
                    return

                if target_category == "New Category...":
                    new_category, ok = QInputDialog.getText(
                        self.window,
                        "New Category",
                        "Enter a new category name:"
                    )
                    if not ok or not new_category.strip():
                        return

                    target_category = new_category.strip()

                    if target_category not in self.window.config.sections():
                        self.window.config.add_section(target_category)

            title = (data.get("title") or "Imported Snippet").strip() or "Imported Snippet"
            content = data.get("content", "")
            title = self.window.make_unique_note_title(target_category, title)

            index = self.window.get_next_item_index(target_category)
            self.window.config[target_category][f"item{index}_title"] = title
            self.window.config[target_category][f"item{index}_content"] = content

            self.window.config_manager.save()
            self.window.load_categories()
            self.window.select_category_by_name(target_category)
            self.window.load_notes()
            self.window.select_note_by_title(title)
            self.window.statusBar().showMessage("Snippet imported.", 3000)
            return

        QMessageBox.warning(self.window, self.app_name, "That file is not a valid Quick Snippet import file.")