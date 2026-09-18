"""Панель со списком преподавателей."""

from __future__ import annotations

from collections.abc import Sequence

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QGroupBox,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from evaluation_app.models import Teacher
from evaluation_app.utils.file_utils import format_score


class TeacherListWidget(QGroupBox):
    """Список преподавателей проекта с кнопками управления."""

    teacherSelected = pyqtSignal(str)
    addRequested = pyqtSignal()
    renameRequested = pyqtSignal(str)
    removeRequested = pyqtSignal(str)
    importRequested = pyqtSignal(str)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__("Преподаватели", parent)
        self._teachers: list[Teacher] = []

        layout = QVBoxLayout(self)

        self.list_widget = QListWidget(self)
        self.list_widget.setAlternatingRowColors(True)
        self.list_widget.currentItemChanged.connect(self._on_current_changed)
        layout.addWidget(self.list_widget, 1)

        self.info_label = QLabel("Преподаватели не добавлены", self)
        self.info_label.setWordWrap(True)
        self.info_label.setStyleSheet("color: #555;")
        layout.addWidget(self.info_label)

        self.add_button = QPushButton("Добавить", self)
        self.rename_button = QPushButton("Переименовать", self)
        self.import_button = QPushButton("Обновить из Excel", self)
        self.remove_button = QPushButton("Удалить", self)

        layout.addWidget(self.add_button)
        layout.addWidget(self.rename_button)
        layout.addWidget(self.import_button)
        layout.addWidget(self.remove_button)

        self.add_button.clicked.connect(self.addRequested.emit)
        self.rename_button.clicked.connect(lambda: self.renameRequested.emit(self.current_teacher_id()))
        self.import_button.clicked.connect(lambda: self.importRequested.emit(self.current_teacher_id()))
        self.remove_button.clicked.connect(lambda: self.removeRequested.emit(self.current_teacher_id()))

        self.set_teachers([], None)

    # --- API -----------------------------------------------------------------

    def set_teachers(self, teachers: Sequence[Teacher], selected_id: str | None = None) -> None:
        """Перестраивает список, по возможности сохраняя выделение."""
        if selected_id is None:
            selected_id = self.current_teacher_id()

        self._teachers = list(teachers)
        self.list_widget.blockSignals(True)
        self.list_widget.clear()

        selected_row = -1
        for row, teacher in enumerate(self._teachers):
            item = QListWidgetItem(teacher.name or "Без имени")
            item.setData(Qt.UserRole, teacher.id)
            details = (
                f"Направлений: {len(teacher.directions)}\n"
                f"Учеников: {len(teacher.student_names())}\n"
                f"Сумма баллов: {format_score(teacher.total_score())}"
            )
            item.setToolTip(details)
            self.list_widget.addItem(item)
            if selected_id and teacher.id == selected_id:
                selected_row = row

        self.list_widget.blockSignals(False)
        if selected_row < 0 and self._teachers:
            selected_row = 0
        if selected_row >= 0:
            self.list_widget.setCurrentRow(selected_row)

        self._update_buttons()
        self._update_info()
        self.teacherSelected.emit(self.current_teacher_id())

    def current_teacher_id(self) -> str:
        """Идентификатор выбранного преподавателя или пустая строка."""
        item = self.list_widget.currentItem()
        if item is None:
            return ""
        return str(item.data(Qt.UserRole) or "")

    def select_teacher(self, teacher_id: str) -> None:
        """Выбирает преподавателя по идентификатору."""
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            if str(item.data(Qt.UserRole) or "") == teacher_id:
                self.list_widget.setCurrentRow(row)
                return

    # --- Внутренние слоты ----------------------------------------------------

    def _on_current_changed(self, current: QListWidgetItem | None, _previous=None) -> None:
        self._update_buttons()
        self._update_info()
        teacher_id = str(current.data(Qt.UserRole) or "") if current is not None else ""
        self.teacherSelected.emit(teacher_id)

    def _update_buttons(self) -> None:
        has_selection = bool(self.current_teacher_id())
        self.rename_button.setEnabled(has_selection)
        self.import_button.setEnabled(has_selection)
        self.remove_button.setEnabled(has_selection)

    def _update_info(self) -> None:
        if not self._teachers:
            self.info_label.setText("Преподаватели не добавлены")
            return

        teacher_id = self.current_teacher_id()
        teacher = next((t for t in self._teachers if t.id == teacher_id), None)
        if teacher is None:
            self.info_label.setText(f"Всего преподавателей: {len(self._teachers)}")
            return

        directions = ", ".join(d.name for d in teacher.directions) or "нет направлений"
        self.info_label.setText(
            f"Всего: {len(self._teachers)} • выбрано: {teacher.name}\n"
            f"Направления: {directions}\n"
            f"Учеников: {len(teacher.student_names())} • "
            f"Сумма баллов: {format_score(teacher.total_score())}"
        )
