"""Виджет выбора сортировки сводной таблицы."""

from __future__ import annotations

from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QComboBox, QHBoxLayout, QLabel, QToolButton, QWidget

from evaluation_app.services.data_sorter import SORT_KEYS, DataSorter


class SortWidget(QWidget):
    """Комбобокс с ключом сортировки и кнопка направления сортировки."""

    sortChanged = pyqtSignal(str, bool)

    def __init__(self, parent: QWidget | None = None, key: str = "score", reverse: bool = True) -> None:
        super().__init__(parent)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        layout.addWidget(QLabel("Сортировка:", self))

        self.combo = QComboBox(self)
        for sort_key, label in SORT_KEYS.items():
            self.combo.addItem(label, sort_key)
        layout.addWidget(self.combo)

        self.order_button = QToolButton(self)
        self.order_button.setCheckable(True)
        self.order_button.setToolTip("Направление сортировки")
        layout.addWidget(self.order_button)

        self._reverse = reverse
        index = self.combo.findData(key)
        if index >= 0:
            self.combo.setCurrentIndex(index)
        self.order_button.setChecked(reverse)

        self.combo.currentIndexChanged.connect(self._on_key_changed)
        self.order_button.toggled.connect(self._on_order_toggled)
        self._update_order_text()

    # --- API -----------------------------------------------------------------

    def current_key(self) -> str:
        """Текущий ключ сортировки."""
        return str(self.combo.currentData() or "score")

    def current_reverse(self) -> bool:
        """Направление сортировки (True — по убыванию)."""
        return self._reverse

    def set_sort(self, key: str, reverse: bool | None = None) -> None:
        """Устанавливает сортировку без лишних сигналов."""
        if reverse is None:
            reverse = DataSorter.default_reverse(key)

        self.blockSignals(True)
        self.combo.blockSignals(True)
        self.order_button.blockSignals(True)
        index = self.combo.findData(key)
        if index >= 0:
            self.combo.setCurrentIndex(index)
        self._reverse = bool(reverse)
        self.order_button.setChecked(self._reverse)
        self._update_order_text()
        self.order_button.blockSignals(False)
        self.combo.blockSignals(False)
        self.blockSignals(False)

    # --- Внутреннее ----------------------------------------------------------

    def _update_order_text(self) -> None:
        self.order_button.setText("↓ убыв." if self._reverse else "↑ возр.")

    def _on_key_changed(self) -> None:
        """При смене ключа используем направление сортировки по умолчанию."""
        self._reverse = DataSorter.default_reverse(self.current_key())
        self.order_button.blockSignals(True)
        self.order_button.setChecked(self._reverse)
        self.order_button.blockSignals(False)
        self._update_order_text()
        self._emit_changed()

    def _on_order_toggled(self, checked: bool) -> None:
        self._reverse = bool(checked)
        self._update_order_text()
        self._emit_changed()

    def _emit_changed(self) -> None:
        self.sortChanged.emit(self.current_key(), self._reverse)
