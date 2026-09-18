"""Дымовые тесты интерфейса (запускаются в offscreen-режиме Qt)."""

from __future__ import annotations

import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

pytest.importorskip("PyQt5")

from PyQt5.QtWidgets import QApplication

from evaluation_app.models import Project
from evaluation_app.services.data_merger import DataMerger
from evaluation_app.services.excel_parser import ExcelParser, SheetParseResult
from evaluation_app.ui.main_window import MainWindow
from evaluation_app.ui.summary_table_view import SummaryTableView
from evaluation_app.ui.teacher_list_widget import TeacherListWidget
from evaluation_app.ui.upload_dialog import SKIP_TEACHER_LABEL, UploadPreviewDialog, UploadSheet
from tests.helpers import make_workbook, sample_project

STANDARD_ROWS = [
    ["ФИО", "Название конкурса", "Итого"],
    ["Иванов Иван", "Олимпиада", 85],
]


@pytest.fixture(scope="module")
def qapp():
    try:
        app = QApplication.instance() or QApplication([])
    except Exception as exc:  # pragma: no cover - зависит от окружения
        pytest.skip(f"Qt недоступен: {exc}")
    yield app


@pytest.fixture()
def window(qapp):
    main_window = MainWindow()
    yield main_window
    # Снимаем флаг изменений, чтобы закрытие не открывало модальный диалог.
    main_window._dirty = False
    main_window.close()


def test_window_uses_app_icon(window) -> None:
    """Иконка приложения попала и в окно, и в набор размеров для панели задач."""
    icon = window.windowIcon()
    assert not icon.isNull()
    assert {16, 32, 48, 256} <= {size.width() for size in icon.availableSizes()}


def test_app_icon_without_file_falls_back(qapp, monkeypatch) -> None:
    """Если файла иконки нет, приложение получает пустую иконку вместо ошибки."""
    from pathlib import Path

    from evaluation_app.ui import app_icon as icon_module

    monkeypatch.setattr(icon_module, "ICONS_SUBPATH", Path("нет-такого-каталога"))
    assert icon_module.app_icon().isNull()


def test_icon_path_inside_frozen_bundle(qapp, tmp_path, monkeypatch) -> None:
    """В собранном exe ресурсы лежат в sys._MEIPASS, а не рядом с модулем."""
    import shutil
    import sys
    from pathlib import Path

    from evaluation_app.ui import app_icon as icon_module

    icons = tmp_path / "evaluation_app" / "resources" / "icons"
    icons.mkdir(parents=True)
    source = Path(__file__).resolve().parent.parent / "evaluation_app" / "resources" / "icons" / "app.ico"
    shutil.copyfile(source, icons / "app.ico")
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)

    assert icon_module.icon_file().parent == icons
    assert not icon_module.app_icon().isNull()


def test_main_window_starts_empty(window) -> None:
    assert window.project.is_empty()
    assert window.summary_view.table.rowCount() == 0
    assert window.teacher_list.list_widget.count() == 0
    assert "Преподавателей: 0" in window.status_label.text()


def test_teacher_list_and_summary_refresh(window) -> None:
    window.project = sample_project()
    window._refresh_all()

    assert window.teacher_list.list_widget.count() == 2
    # По умолчанию сортировка по баллам по убыванию: сначала «Иванов Иван» (125).
    assert window.summary_view.table.rowCount() == 3
    assert window.summary_view.table.item(0, 0).text() == "Иванов Иван"
    assert window.summary_view.table.item(0, 1).text() == "125"

    # Выбор преподавателя обновляет детальное представление.
    window.teacher_list.select_teacher(window.project.teachers[0].id)
    assert window.direction_view.title_label.text() == "Преподаватель: Петрова А.И."
    assert window.direction_view.tree.topLevelItemCount() == 1


def test_search_and_sorting_signals(window) -> None:
    window.project = sample_project()
    window._refresh_all()

    window.summary_view.set_search_text("кузнец")
    assert window.summary_view.table.rowCount() == 1
    assert window.summary_view.table.item(0, 0).text() == "Кузнецов Лев"
    assert "из 3" in window.summary_view.count_label.text()

    window.summary_view.set_search_text("")
    window._on_sort_changed("name", False)
    assert window.summary_view.table.item(0, 0).text() == "Иванов Иван"
    assert window.summary_view.table.item(2, 0).text() == "Петрова Анна"

    window._on_sort_changed("score", False)
    assert window.summary_view.table.item(0, 0).text() == "Кузнецов Лев"


MATRIX_ROWS = [
    ["ФИО", "за вход в систему", "УЧЕБНАЯ, ПРОЕКТНАЯ ДЕЯТЕЛЬНОСТЬ (КОНКУРСЫ)", None, "ИТОГО"],
    [None, None, "Участие (внутренний)", "Разовое мероприятие", None],
    [None, 30, 30, "40/50/60", None],
    ["Соколова Наталья Фёдоровна", 10, 10, 0, 10],
]


def test_column_mapping_dialog_preselects_matrix(qapp, tmp_path) -> None:
    """Группа «... (КОНКУРСЫ)» не должна подставляться как название конкурса."""
    from evaluation_app.services.excel_parser import MODE_CRITERIA
    from evaluation_app.ui.upload_dialog import ColumnMappingDialog

    path = make_workbook(tmp_path / "matrix.xlsx", {"Лист": MATRIX_ROWS})
    parser = ExcelParser(str(path))
    parser.load()

    dialog = ColumnMappingDialog(parser, "Лист", None)
    assert dialog.mode() == MODE_CRITERIA
    assert dialog.columns().get("student_name") == 1
    assert "competition_name" not in dialog.columns()
    assert dialog.columns().get("total_score") == 5
    dialog.close()


def test_sort_widget_key_change_resets_direction(qapp) -> None:
    from evaluation_app.ui.sort_widget import SortWidget

    widget = SortWidget(key="name", reverse=False)
    events: list[tuple] = []
    widget.sortChanged.connect(lambda key, reverse: events.append((key, reverse)))

    widget.combo.setCurrentIndex(widget.combo.findData("score"))
    assert events[-1] == ("score", True)  # баллы — по убыванию

    widget.order_button.setChecked(False)
    assert events[-1] == ("score", False)
    assert widget.current_reverse() is False

    widget.set_sort("competitions", True)
    assert (widget.current_key(), widget.current_reverse()) == ("competitions", True)


def test_summary_view_handles_empty_records(qapp) -> None:
    view = SummaryTableView()
    view.set_records([], total_count=5)
    assert view.table.rowCount() == 0
    assert "0 из 5" in view.count_label.text()


def test_teacher_list_widget_signals(qapp) -> None:
    widget = TeacherListWidget()
    received: list[str] = []
    widget.teacherSelected.connect(received.append)

    project = sample_project()
    widget.set_teachers(project.teachers, project.teachers[1].id)
    assert widget.current_teacher_id() == project.teachers[1].id
    assert received[-1] == project.teachers[1].id
    assert "Сидоров П.П." in widget.info_label.text()

    widget.set_teachers([])
    assert widget.current_teacher_id() == ""
    assert widget.info_label.text() == "Преподаватели не добавлены"


def test_upload_dialog_selection(qapp, tmp_path) -> None:
    path = make_workbook(tmp_path / "Петрова_2024.xlsx", {"Математика": STANDARD_ROWS, "Физика": STANDARD_ROWS})
    parser = ExcelParser(str(path))
    items = []
    for sheet_name in parser.sheet_names:
        items.append(
            UploadSheet(
                file_path=str(path),
                sheet_name=sheet_name,
                parser=parser,
                default_teacher=parser.default_teacher_name(),
                result=SheetParseResult(sheet_name=sheet_name, direction=parser.parse_sheet(sheet_name).direction),
            )
        )

    dialog = UploadPreviewDialog(items, [], None)
    assert dialog.mapping_table.rowCount() == 2
    assert dialog.get_selection() == [
        (str(path), "Математика", "Петрова"),
        (str(path), "Физика", "Петрова"),
    ]

    # Отмечаем второй лист как «пропустить».
    combo = dialog.mapping_table.cellWidget(1, 2)
    combo.setCurrentIndex(combo.findText(SKIP_TEACHER_LABEL))
    assert dialog.get_selection() == [(str(path), "Математика", "Петрова")]

    # Другому преподавателю можно назначить свой лист.
    combo.setCurrentText("Сидоров П.П.")
    selection = dialog.get_selection()
    assert selection[0][2] == "Петрова"
    assert selection[1][2] == "Сидоров П.П."
    dialog.close()


def test_main_window_undo_and_teacher_management(window) -> None:
    window.project = Project(project_name="П")
    window._push_history()
    window.project.add_teacher(name="Первый")
    window._refresh_all()
    assert window.teacher_list.list_widget.count() == 1

    window._on_undo()
    assert window.project.is_empty()
    window._on_undo()  # повторная отмена без истории не падает
    assert window.project.is_empty()


def test_main_window_handler_guards(window) -> None:
    # Обработчики без выбранного преподавателя не должны падать (диалоги не открываются).
    window.project = Project()
    assert window._visible_records() == []
    window._on_teacher_selected("")
    assert window.direction_view.teacher() is None
    assert window._apply_import_selection([]) == (0, 0, [])


def test_import_selection_applies_to_project(window, tmp_path) -> None:
    path = make_workbook(tmp_path / "Петрова.xlsx", {"Математика": STANDARD_ROWS})
    added, replaced, warnings = window._apply_import_selection(
        [(str(path), "Математика", "Петрова А.И.")]
    )
    assert (added, replaced) == (1, 0)
    assert warnings == []
    assert window.project.teachers[0].total_score() == 85

    # Повторный импорт того же листа обновляет данные, а не дублирует их.
    added, replaced, _ = window._apply_import_selection(
        [(str(path), "Математика", "Петрова А.И.")]
    )
    assert (added, replaced) == (0, 1)
    assert len(window.project.teachers) == 1
    assert window.project.teachers[0].total_score() == 85

    # Ошибочный лист попадает в предупреждения.
    bad = make_workbook(tmp_path / "Плохой.xlsx", {"Лист": [["Нет", "Нужных", "Столбцов"]]})
    _added, _replaced, warnings = window._apply_import_selection([(str(bad), "Лист", "Плохой")])
    assert warnings


def test_project_statistics_status(window) -> None:
    window.project = sample_project()
    window._refresh_all()
    stats = DataMerger.build_statistics(window.project)
    assert stats["students"] == 3
    assert "Учеников: 3" in window.status_label.text()


def test_teacher_without_data_selection(window) -> None:
    project = Project()
    project.add_teacher(name="Пустой")
    window.project = project
    window._refresh_all()
    window.teacher_list.select_teacher(project.teachers[0].id)
    assert window.direction_view.tree.topLevelItemCount() == 0
    assert window.direction_view.teacher() is project.teachers[0]


def test_create_template_and_import_dialog_flow(window, tmp_path) -> None:
    """Шаблон, созданный приложением, можно сразу импортировать."""
    from evaluation_app.services.excel_export import ExcelExporter

    template = tmp_path / "Шаблон.xlsx"
    ExcelExporter.create_template(str(template), sheet_names=("Направление 1",))

    items = window._collect_upload_items([str(template)])
    assert len(items) == 1
    assert items[0].rows_count == len(ExcelExporter.template_evaluations())

    added, _replaced, _warnings = window._apply_import_selection(
        [(str(template), "Направление 1", "Тестовый преподаватель")]
    )
    assert added == 1
    assert window.project.find_teacher("Тестовый преподаватель") is not None
