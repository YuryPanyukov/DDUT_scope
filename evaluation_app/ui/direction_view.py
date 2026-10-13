"""Детальный просмотр данных выбранного преподавателя.

Данные показываются матрицей, как во входных оценочных листах: строки —
ученики направления, столбцы — критерии (конкурсы), последний столбец —
итоговая сумма баллов ученика.
"""

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

from evaluation_app.models import Direction, Teacher
from evaluation_app.services.data_merger import DataMerger
from evaluation_app.ui.column_controls import enable_column_controls, make_columns_interactive
from evaluation_app.utils.file_utils import format_score
from evaluation_app.utils.validators import name_key, normalize_header, normalize_name

# Заголовки без критериев (пустой преподаватель или данные без оценок).
BASE_HEADERS = ("Направление / ФИО", "Итого баллов")


class DirectionView(QWidget):
    """Матрица «ученики × критерии» для выбранного преподавателя."""

    exportRequested = pyqtSignal(str)  # id преподавателя

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._teacher: Teacher | None = None
        # Сортировка учеников внутри направления: "name" | "total" | None.
        self._sort_key: str | None = None
        self._sort_reverse = False
        # Скрытые пользователем столбцы: хранятся по названию, потому что
        # матрица пересобирается (смена преподавателя, сортировка).
        self._hidden_columns: set[str] = set()
        self._headers: list[str] = []

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
        self.tree.setAlternatingRowColors(True)
        self.tree.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tree.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tree.setRootIsDecorated(True)
        layout.addWidget(self.tree, 1)

        header = self.tree.header()
        header.setSectionsClickable(True)
        header.sectionClicked.connect(self._on_header_clicked)
        enable_column_controls(self.tree, on_columns_changed=self._remember_hidden_columns)

        self.export_button.clicked.connect(self._on_export_clicked)

    # --- API -----------------------------------------------------------------

    def show_teacher(self, teacher: Teacher | None) -> None:
        """Отображает данные преподавателя матрицей «ученики × критерии»."""
        self._teacher = teacher
        self.tree.clear()

        if teacher is None:
            self._apply_headers([])
            self.tree.header().setSortIndicatorShown(False)
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

        criteria = self._criteria_columns(teacher)
        self._apply_headers(criteria)
        criterion_index = {normalize_header(name): pos for pos, name in enumerate(criteria)}

        for direction in teacher.directions:
            direction_item = QTreeWidgetItem(
                [
                    f"{direction.name} ({len(direction.evaluations)})",
                    *[""] * len(criteria),
                    format_score(direction.total_score()),
                ]
            )
            self._right_align(direction_item)
            font = direction_item.font(0)
            font.setBold(True)
            direction_item.setFont(0, font)

            for student_name, cell_texts, student_total in self._student_rows(
                direction, criterion_index, len(criteria)
            ):
                child = QTreeWidgetItem([student_name, *cell_texts, format_score(student_total)])
                self._right_align(child)
                direction_item.addChild(child)

            self.tree.addTopLevelItem(direction_item)

        self.tree.expandToDepth(0)
        self._update_sort_indicator()

    def clear(self) -> None:
        """Очищает представление."""
        self.show_teacher(None)

    def teacher(self) -> Teacher | None:
        """Текущий отображаемый преподаватель."""
        return self._teacher

    # --- Внутреннее ----------------------------------------------------------

    def _apply_headers(self, criteria: list[str]) -> None:
        """Настраивает столбцы матрицы: ФИО | критерии | итоговая сумма."""
        headers = [BASE_HEADERS[0], *criteria, BASE_HEADERS[1]]
        if headers == self._headers:
            # Структура не менялась — ширины и скрытые столбцы пользователя
            # сохраняются (иначе сортировка сбрасывала бы размеры).
            return
        self._headers = headers

        self.tree.setColumnCount(len(headers))
        self.tree.setHeaderLabels(headers)

        header = self.tree.header()
        for index in range(len(headers)):
            mode = QHeaderView.Stretch if index == 0 else QHeaderView.ResizeToContents
            header.setSectionResizeMode(index, mode)

        for index, label in enumerate(headers):
            if normalize_header(label) in self._hidden_columns:
                self.tree.hideColumn(index)

        if self.isVisible():
            # Геометрия уже разложена — ширины по содержимому и ручной режим.
            make_columns_interactive(self.tree)

    def _remember_hidden_columns(self) -> None:
        """Запоминает скрытые столбцы по названию — их вернёт пересборка матрицы."""
        header_item = self.tree.headerItem()
        if header_item is None:
            return
        self._hidden_columns = {
            normalize_header(header_item.text(index))
            for index in range(self.tree.columnCount())
            if self.tree.isColumnHidden(index)
        }

    def _student_rows(
        self,
        direction: Direction,
        criterion_index: dict[str, int],
        criteria_count: int,
    ) -> list[tuple[str, list[str], float]]:
        """Строки матрицы направления: (ФИО, ячейки критериев, итоговая сумма)."""
        rows: list[tuple[str, list[str], float]] = []
        for student_name, evaluations in DataMerger.group_by_student(direction.evaluations):
            cells = [0.0] * criteria_count
            filled = [False] * criteria_count
            for evaluation in evaluations:
                position = criterion_index.get(normalize_header(evaluation.competition_name))
                if position is None:
                    continue
                filled[position] = True
                cells[position] += float(evaluation.total_score or 0.0)
            student_total = float(sum(evaluation.total_score or 0.0 for evaluation in evaluations))
            cell_texts = [
                format_score(cells[position]) if filled[position] else ""
                for position in range(criteria_count)
            ]
            rows.append((student_name, cell_texts, student_total))

        # Сортировка только по ФИО или итогу; по критериям — нет (ячейки текст,
        # порядок получился бы ложным, «100» < «25»).
        if self._sort_key == "name":
            rows.sort(key=lambda row: name_key(row[0]), reverse=self._sort_reverse)
        elif self._sort_key == "total":
            rows.sort(key=lambda row: row[2], reverse=self._sort_reverse)
        return rows

    @staticmethod
    def _criteria_columns(teacher: Teacher) -> list[str]:
        """Уникальные названия критериев во всех направлениях (первое вхождение)."""
        criteria: list[str] = []
        seen: set[str] = set()
        for direction in teacher.directions:
            for evaluation in direction.evaluations:
                label = normalize_name(evaluation.competition_name)
                if not label:
                    continue
                key = normalize_header(label)
                if key in seen:
                    continue
                seen.add(key)
                criteria.append(label)
        return criteria

    @staticmethod
    def _right_align(item: QTreeWidgetItem) -> None:
        """Выравнивает числовые столбцы матрицы вправо."""
        for column in range(1, item.columnCount()):
            item.setTextAlignment(column, Qt.AlignRight | Qt.AlignVCenter)

    def _on_header_clicked(self, section: int) -> None:
        """Сортирует учеников по клику: 0 — ФИО, последний столбец — итог."""
        if self._teacher is None:
            return

        last_section = self.tree.columnCount() - 1
        if section == 0:
            key = "name"
        elif section == last_section:
            key = "total"
        else:
            return

        if self._sort_key == key:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_key = key
            # Баллы — по убыванию, как в сводной таблице; ФИО — по возрастанию.
            self._sort_reverse = key == "total"

        self.show_teacher(self._teacher)

    def _update_sort_indicator(self) -> None:
        """Показывает стрелку сортировки на активном столбце заголовка."""
        header = self.tree.header()
        if self._sort_key is None:
            header.setSortIndicatorShown(False)
            return
        section = 0 if self._sort_key == "name" else self.tree.columnCount() - 1
        order = Qt.DescendingOrder if self._sort_reverse else Qt.AscendingOrder
        header.setSortIndicatorShown(True)
        header.setSortIndicator(section, order)

    def _on_export_clicked(self) -> None:
        if self._teacher is not None:
            self.exportRequested.emit(self._teacher.id)
