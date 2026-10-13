"""Управление столбцами таблиц проекта: ширина, порядок, скрытие.

Столбцы любой таблицы можно перетаскивать за заголовок и менять им ширину
перетаскиванием границы. Правый клик по названию столбца открывает меню:

* «Скрыть столбец …» — скрывает столбец, на котором вызвано меню;
* «Показать все скрытые столбцы» — возвращает все ранее скрытые.

Ширины по содержимому считаются один раз — при первом показе вида: до ``show``
геометрия не разложена, и столбцы получили бы ширину по ещё не показанному
окну. Дальше столбцы остаются в ручном режиме (``Interactive``), поэтому
пользователь может и двигать их, и изменять ширину.
"""

from __future__ import annotations

from collections.abc import Callable

from PyQt5.QtCore import QEvent, QObject, Qt
from PyQt5.QtWidgets import QAbstractItemView, QHeaderView, QMenu

# Тексты пунктов меню — часть API, их же видит пользователь.
HIDE_COLUMN_TEXT = "Скрыть столбец"
SHOW_ALL_COLUMNS_TEXT = "Показать все скрытые столбцы"

ChangedCallback = Callable[[], None] | None


def header_of(view: QAbstractItemView) -> QHeaderView:
    """Заголовок горизонтальных столбцов: у таблицы — horizontalHeader(), у дерева — header()."""
    horizontal = getattr(view, "horizontalHeader", None)
    if horizontal is not None:
        return horizontal()
    return view.header()


def enable_column_controls(
    view: QAbstractItemView, on_columns_changed: ChangedCallback = None
) -> None:
    """Включает перетаскивание столбцов, изменение их ширины и меню ПКМ.

    :param on_columns_changed: вызывается после скрытия/показа столбцов —
        вид может запомнить выбор пользователя (матрица преподавателя
        пересобирается при смене преподавателя и сортировке).
    """
    header = header_of(view)
    header.setSectionsMovable(True)
    header.setContextMenuPolicy(Qt.CustomContextMenu)
    header.customContextMenuRequested.connect(
        lambda position: _open_header_menu(view, position, on_columns_changed)
    )
    watcher = _InteractiveOnShow(view)
    view.installEventFilter(watcher)


def make_columns_interactive(view: QAbstractItemView) -> None:
    """Считает ширины по текущим режимам и переводит столбцы в ручной режим."""
    header = header_of(view)
    # header.resizeSections() у PyQt5 — защищённый метод, и на заголовке,
    # созданном самим видом (не из Python), он запрещён (виден как RuntimeError
    # «no access to protected functions» в Linux); публичные слоты вида делают
    # то же самое для любого типа: таблицы — одним вызовом, деревья — по столбцу.
    columns_resize = getattr(view, "resizeColumnsToContents", None)
    if columns_resize is not None:
        columns_resize()
    else:
        for index in range(header.count()):
            view.resizeColumnToContents(index)
    for index in range(header.count()):
        header.setSectionResizeMode(index, QHeaderView.Interactive)


def build_header_menu(
    view: QAbstractItemView,
    logical_index: int,
    on_columns_changed: ChangedCallback = None,
) -> QMenu:
    """Собирает меню заголовка: скрыть столбец / показать все скрытые.

    Пункты уже связаны с действием, поэтому тесты могут вызвать
    ``action.trigger()`` без показа меню на экране.
    """
    menu = QMenu(view)
    _fill_header_menu(menu, view, logical_index, on_columns_changed)
    return menu


def _fill_header_menu(
    menu: QMenu,
    view: QAbstractItemView,
    logical_index: int,
    on_columns_changed: ChangedCallback,
) -> None:
    label = _header_label(view, logical_index)
    title = f"{HIDE_COLUMN_TEXT} «{label}»" if label else HIDE_COLUMN_TEXT
    hide_action = menu.addAction(title)
    hide_action.setEnabled(
        0 <= logical_index < view.columnCount()
        and not view.isColumnHidden(logical_index)
        and _visible_count(view) > 1
    )
    hide_action.triggered.connect(
        lambda _checked=False, index=logical_index: _hide_column(view, index, on_columns_changed)
    )

    menu.addSeparator()

    show_all_action = menu.addAction(SHOW_ALL_COLUMNS_TEXT)
    show_all_action.setEnabled(_hidden_count(view) > 0)
    show_all_action.triggered.connect(
        lambda _checked=False: _show_all_columns(view, on_columns_changed)
    )


# --- Внутреннее -------------------------------------------------------------

class _InteractiveOnShow(QObject):
    """Переводит столбцы в ручной режим при показе вида (первый раз или повторно).

    Живёт вместе с видом: фильтр устанавливается на сам вид, поэтому
    объект-фильтр удерживается Qt как его дочерний.
    """

    def __init__(self, view: QAbstractItemView) -> None:
        super().__init__(view)
        self._view = view

    def eventFilter(self, _watched: QObject, event: QEvent) -> bool:
        if event.type() == QEvent.Show and not _fully_interactive(self._view):
            make_columns_interactive(self._view)
        return False


def _fully_interactive(view: QAbstractItemView) -> bool:
    header = header_of(view)
    return all(
        header.sectionResizeMode(index) == QHeaderView.Interactive
        for index in range(header.count())
    )


def _open_header_menu(
    view: QAbstractItemView, position, on_columns_changed: ChangedCallback
) -> None:
    """Собирает и показывает меню под курсором (ПКМ по названию столбца)."""
    header = header_of(view)
    previous = getattr(view, "_column_controls_menu", None)
    if previous is not None:
        previous.hide()
        previous.deleteLater()

    menu = build_header_menu(view, header.logicalIndexAt(position), on_columns_changed)
    # Ссылка хранится на виде — popup-меню должно пережить текущий вызов,
    # а предыдущее (если висело) подчищается выше.
    view._column_controls_menu = menu
    menu.popup(header.mapToGlobal(position))


def _hide_column(view: QAbstractItemView, index: int, on_columns_changed: ChangedCallback) -> None:
    if not (0 <= index < view.columnCount()):
        return
    if view.isColumnHidden(index) or _visible_count(view) <= 1:
        return
    view.hideColumn(index)
    _notify(on_columns_changed)


def _show_all_columns(view: QAbstractItemView, on_columns_changed: ChangedCallback) -> None:
    for index in range(view.columnCount()):
        if view.isColumnHidden(index):
            view.showColumn(index)
    _notify(on_columns_changed)


def _visible_count(view: QAbstractItemView) -> int:
    return sum(1 for index in range(view.columnCount()) if not view.isColumnHidden(index))


def _hidden_count(view: QAbstractItemView) -> int:
    return view.columnCount() - _visible_count(view)


def _header_label(view: QAbstractItemView, logical_index: int) -> str:
    """Название столбца из шапки таблицы или дерева."""
    if logical_index < 0:
        return ""
    horizontal_item = getattr(view, "horizontalHeaderItem", None)
    if horizontal_item is not None:
        item = horizontal_item(logical_index)
        return item.text() if item is not None else ""
    header_item = view.headerItem()  # QTreeWidget
    if header_item is None:
        return ""
    if logical_index >= header_item.columnCount():
        return ""
    return header_item.text(logical_index)


def _notify(on_columns_changed: ChangedCallback) -> None:
    if on_columns_changed is not None:
        on_columns_changed()
