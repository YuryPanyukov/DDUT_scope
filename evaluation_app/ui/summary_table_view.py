"""Сводная таблица результатов (ФИО → итого баллов)."""

from __future__ import annotations

from collections.abc import Sequence

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from evaluation_app.ui.column_controls import enable_column_controls
from evaluation_app.ui.sort_widget import SortWidget
from evaluation_app.utils.file_utils import format_score

COLUMNS = ("ФИО", "Итого баллов", "Оценок", "Преподаватели", "Направления", "Конкурсы")
# Индекс столбца → ключ сортировки (для клика по заголовку).
COLUMN_SORT_KEYS = {0: "name", 1: "score", 2: "entries"}


class SummaryTableView(QWidget):
    """Сводная таблица с поиском, сортировкой и экспортом."""

    sortChanged = pyqtSignal(str, bool)
    searchChanged = pyqtSignal(str)
    exportRequested = pyqtSignal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._records: list[dict] = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        title_row = QHBoxLayout()
        self.title_label = QLabel("Сводная таблица", self)
        self.title_label.setStyleSheet("font-weight: bold;")
        title_row.addWidget(self.title_label)
        title_row.addStretch(1)

        self.sort_widget = SortWidget(self)
        title_row.addWidget(self.sort_widget)

        self.export_button = QPushButton("Экспорт в Excel", self)
        title_row.addWidget(self.export_button)
        layout.addLayout(title_row)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Поиск по ФИО:", self))
        self.search_edit = QLineEdit(self)
        self.search_edit.setPlaceholderText("Начните вводить фамилию ученика...")
        self.search_edit.setClearButtonEnabled(True)
        search_row.addWidget(self.search_edit, 1)
        self.count_label = QLabel("Записей: 0", self)
        search_row.addWidget(self.count_label)
        layout.addLayout(search_row)

        self.table = QTableWidget(self)
        self.table.setColumnCount(len(COLUMNS))
        self.table.setHorizontalHeaderLabels(list(COLUMNS))
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(False)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Interactive)
        header.setSectionResizeMode(4, QHeaderView.Interactive)
        header.setSectionResizeMode(5, QHeaderView.Stretch)
        header.setSectionsClickable(True)
        header.sectionClicked.connect(self._on_header_clicked)
        layout.addWidget(self.table, 1)

        # Столбцы можно двигать, менять им ширину и скрывать (ПКМ по заголовку).
        enable_column_controls(self.table)

        self.sort_widget.sortChanged.connect(self.sortChanged.emit)
        self.search_edit.textChanged.connect(self.searchChanged.emit)
        self.export_button.clicked.connect(self.exportRequested.emit)

    # --- API -----------------------------------------------------------------

    def set_records(self, records: Sequence[dict], total_count: int | None = None) -> None:
        """Заполняет таблицу записями."""
        self._records = list(records)
        total = len(self._records) if total_count is None else total_count

        self.table.setRowCount(len(self._records))
        for row, record in enumerate(self._records):
            self.table.setItem(row, 0, QTableWidgetItem(str(record.get("student_name", ""))))

            score_item = QTableWidgetItem(format_score(record.get("total_score")))
            score_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            score_item.setData(Qt.UserRole, float(record.get("total_score") or 0.0))
            self.table.setItem(row, 1, score_item)

            entries_item = QTableWidgetItem(str(int(record.get("entries") or 0)))
            entries_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(row, 2, entries_item)

            self.table.setItem(row, 3, QTableWidgetItem(", ".join(record.get("teachers_involved") or [])))
            self.table.setItem(row, 4, QTableWidgetItem(", ".join(record.get("directions") or [])))
            self.table.setItem(row, 5, QTableWidgetItem(", ".join(record.get("competitions") or [])))

        if total and total != len(self._records):
            self.count_label.setText(f"Показано: {len(self._records)} из {total}")
            self.title_label.setText("Сводная таблица (поиск)")
        else:
            self.count_label.setText(f"Записей: {len(self._records)}")
            self.title_label.setText("Сводная таблица")

    def set_search_text(self, text: str) -> None:
        """Устанавливает текст поиска."""
        self.search_edit.setText(text)

    # --- Внутреннее ----------------------------------------------------------

    def _on_header_clicked(self, section: int) -> None:
        key = COLUMN_SORT_KEYS.get(section)
        if key is None:
            return
        reverse = not self.sort_widget.current_reverse() if self.sort_widget.current_key() == key else False
        self.sort_widget.set_sort(key, reverse)
        self.sortChanged.emit(key, reverse)
