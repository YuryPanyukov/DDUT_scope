"""Детальный просмотр данных выбранного преподавателя."""

from __future__ import annotations

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from evaluation_app.models import Teacher
from evaluation_app.utils.file_utils import format_score


class DirectionView(QWidget):
    """Дерево «направление → оценки» для выбранного преподавателя."""

    exportRequested = pyqtSignal(str)  # id преподавателя

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._teacher: Teacher | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        header_row = QHBoxLayout()
        self.title_label = QLabel("Преподаватель не выбран", self)
        self.title_label.setStyleSheet("font-weight: bold;")
        header_row.addWidget(self.title_label)
        header_row.addStretch(1)
        self.export_button = QPushButton("Экспорт преподавателя", self)
        self.export_button.setEnabled(False)
        header_row.addWidget(self.export_button)
        layout.addLayout(header_row)

        self.summary_label = QLabel("Выберите преподавателя в списке слева.", self)
        self.summary_label.setWordWrap(True)
        self.summary_label.setStyleSheet("color: #555;")
        layout.addWidget(self.summary_label)

        self.tree = QTreeWidget(self)
        self.tree.setColumnCount(3)
        self.tree.setHeaderLabels(["Направление / ФИО", "Конкурс", "Баллы"])
        self.tree.setAlternatingRowColors(True)
        self.tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tree.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tree.setRootIsDecorated(True)
        header = self.tree.header()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        layout.addWidget(self.tree, 1)

        self.export_button.clicked.connect(self._on_export_clicked)

    # --- API -----------------------------------------------------------------

    def show_teacher(self, teacher: Teacher | None) -> None:
        """Отображает данные преподавателя."""
        self._teacher = teacher
        self.tree.clear()

        if teacher is None:
            self.title_label.setText("Преподаватель не выбран")
            self.summary_label.setText("Выберите преподавателя в списке слева.")
            self.export_button.setEnabled(False)
            return

        self.title_label.setText(f"Преподаватель: {teacher.name or 'Без имени'}")
        self.summary_label.setText(
            f"Направлений: {len(teacher.directions)} • "
            f"Учеников: {len(teacher.student_names())} • "
            f"Оценок: {len(teacher.evaluations())} • "
            f"Сумма баллов: {format_score(teacher.total_score())}"
        )
        self.export_button.setEnabled(bool(teacher.directions))

        for direction in teacher.directions:
            direction_item = QTreeWidgetItem(
                [f"{direction.name} ({len(direction.evaluations)})", "", format_score(direction.total_score())]
            )
            direction_item.setTextAlignment(2, Qt.AlignRight | Qt.AlignVCenter)
            font = direction_item.font(0)
            font.setBold(True)
            direction_item.setFont(0, font)
            for evaluation in direction.evaluations:
                child = QTreeWidgetItem(
                    [
                        evaluation.student_name,
                        evaluation.competition_name or "—",
                        format_score(evaluation.total_score),
                    ]
                )
                child.setTextAlignment(2, Qt.AlignRight | Qt.AlignVCenter)
                direction_item.addChild(child)
            self.tree.addTopLevelItem(direction_item)

        self.tree.expandToDepth(0)

    def clear(self) -> None:
        """Очищает представление."""
        self.show_teacher(None)

    def teacher(self) -> Teacher | None:
        """Текущий отображаемый преподаватель."""
        return self._teacher

    # --- Внутреннее ----------------------------------------------------------

    def _on_export_clicked(self) -> None:
        if self._teacher is not None:
            self.exportRequested.emit(self._teacher.id)
