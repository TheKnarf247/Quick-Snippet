from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QLabel, QToolButton, QVBoxLayout, QHBoxLayout, QWidget


class NoteListItemWidget(QWidget):
    def __init__(self, title: str, preview: str, copy_callback, icon_path_func, theme: str):
        super().__init__()

        self.title_label = QLabel(title)
        self.title_label.setObjectName("snippetTitleLabel")
        self.title_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)

        title_font = self.title_label.font()
        title_font.setPointSize(14)
        title_font.setBold(True)
        self.title_label.setFont(title_font)

        preview_text = preview.replace("\n", " ").strip()
        if len(preview_text) > 90:
            preview_text = preview_text[:87] + "..."

        self.preview_label = QLabel(preview_text)
        self.preview_label.setObjectName("snippetPreviewLabel")
        self.preview_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.preview_label.setWordWrap(True)

        preview_font = self.preview_label.font()
        preview_font.setPointSize(11)
        preview_font.setBold(False)
        self.preview_label.setFont(preview_font)

        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(4)
        text_layout.addWidget(self.title_label)
        text_layout.addWidget(self.preview_label)
        text_layout.addStretch(1)

        self.copy_button = QToolButton()
        self.copy_button.setToolTip("Copy snippet")
        self.copy_button.setProperty("icon_filename", "IconCopy.png")
        self.copy_button.setIcon(QIcon(icon_path_func("IconCopy.png", theme)))
        self.copy_button.setIconSize(QSize(24, 24))
        self.copy_button.setFixedSize(40, 40)
        self.copy_button.clicked.connect(copy_callback)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 10, 8)
        layout.setSpacing(10)
        layout.addLayout(text_layout, 1)
        layout.addWidget(self.copy_button, 0, Qt.AlignTop)

        self.setMinimumHeight(72)