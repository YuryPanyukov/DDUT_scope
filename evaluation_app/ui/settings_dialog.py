"""Диалоги настроек проекта и приложения (см. п. 6.2–6.3 плана)."""

from __future__ import annotations

import os
from datetime import date

from PyQt5.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


def default_academic_year(today: date | None = None) -> str:
    """Учебный год по умолчанию (сентябрь открывает новый год)."""
    today = today or date.today()
    start = today.year if today.month >= 8 else today.year - 1
    return f"{start}-{start + 1}"


class ProjectSettingsDialog(QDialog):
    """Диалог создания нового проекта."""

    def __init__(
        self,
        parent: QWidget | None = None,
        project_name: str = "",
        academic_year: str = "",
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Новый проект")
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.project_name_edit = QLineEdit(project_name, self)
        self.project_name_edit.setPlaceholderText("Например: Оценки преподавателей ДДЮТ")
        form.addRow("Название проекта:", self.project_name_edit)

        self.year_edit = QLineEdit(academic_year or default_academic_year(), self)
        form.addRow("Учебный год:", self.year_edit)

        layout.addLayout(form)
        layout.addWidget(QLabel("Эти данные сохраняются вместе с проектом.", self))

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        buttons.button(QDialogButtonBox.Ok).setText("Создать")
        buttons.button(QDialogButtonBox.Cancel).setText("Отмена")
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_accept(self) -> None:
        if not self.project_name_edit.text().strip():
            self.project_name_edit.setText("Новый проект")
        self.accept()

    def get_project_name(self) -> str:
        """Название проекта."""
        return self.project_name_edit.text().strip()

    def get_academic_year(self) -> str:
        """Учебный год."""
        return self.year_edit.text().strip()


class AppSettingsDialog(QDialog):
    """Диалог настроек приложения (каталог сохранения по умолчанию)."""

    def __init__(self, parent: QWidget | None = None, default_save_path: str = "") -> None:
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.setMinimumWidth(520)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        path_row = QHBoxLayout()
        self.path_edit = QLineEdit(default_save_path, self)
        path_row.addWidget(self.path_edit, 1)
        browse_button = QPushButton("Обзор...", self)
        browse_button.clicked.connect(self._on_browse)
        path_row.addWidget(browse_button)

        form.addRow("Каталог сохранения:", path_row)
        layout.addLayout(form)

        hint = QLabel(
            "В этом каталоге по умолчанию будут сохраняться проекты и файлы экспорта.",
            self,
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #555;")
        layout.addWidget(hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        buttons.button(QDialogButtonBox.Ok).setText("Сохранить")
        buttons.button(QDialogButtonBox.Cancel).setText("Отмена")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_browse(self) -> None:
        directory = QFileDialog.getExistingDirectory(
            self, "Выберите каталог", self.path_edit.text() or os.path.expanduser("~")
        )
        if directory:
            self.path_edit.setText(directory)

    def get_default_save_path(self) -> str:
        """Выбранный каталог сохранения."""
        return self.path_edit.text().strip()
