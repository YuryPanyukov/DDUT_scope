"""Тесты экспорта в Excel и создания шаблонов (этап 8 плана)."""

from __future__ import annotations

import openpyxl

from evaluation_app.services.data_merger import DataMerger
from evaluation_app.services.data_sorter import DataSorter
from evaluation_app.services.excel_export import ExcelExporter
from evaluation_app.services.excel_parser import ExcelParser
from evaluation_app.services.project_manager import ProjectManager
from tests.helpers import sample_project


def test_create_template_is_readable_by_parser(tmp_path) -> None:
    path = tmp_path / "Шаблон.xlsx"
    ExcelExporter.create_template(str(path), sheet_names=("Направление 1", "Направление 2"))

    workbook = openpyxl.load_workbook(str(path))
    assert workbook.sheetnames == ["Направление 1", "Направление 2"]
    assert [cell.value for cell in workbook["Направление 1"][1]] == ["ФИО", "Название конкурса", "Итого"]

    parser = ExcelParser(str(path), teacher_name="Тестовый преподаватель")
    result = parser.parse_sheet("Направление 1")
    assert result.evaluations_count == len(ExcelExporter.template_evaluations())

    teachers = parser.parse_all_teachers()
    assert len(teachers) == 1
    assert [d.name for d in teachers[0].directions] == ["Направление 1", "Направление 2"]


def test_create_template_with_custom_rows(tmp_path) -> None:
    path = tmp_path / "Пустой шаблон.xlsx"
    ExcelExporter.create_template(str(path), sheet_names=("Математика",), rows=[])
    parser = ExcelParser(str(path), teacher_name="Т")
    result = parser.parse_sheet("Математика")
    assert result.evaluations_count == 0


def test_export_summary(tmp_path) -> None:
    project = sample_project()
    records = DataSorter.sort_by_score(DataMerger.merge_by_student(project))
    path = tmp_path / "Сводная.xlsx"

    ExcelExporter.export_summary(records, str(path), project=project)

    workbook = openpyxl.load_workbook(str(path))
    worksheet = workbook.active
    assert worksheet.cell(row=1, column=1).value == "Тестовый проект (2025-2026)"
    assert worksheet.cell(row=3, column=1).value == "ФИО"
    assert worksheet.cell(row=4, column=1).value == "Иванов Иван"
    assert worksheet.cell(row=4, column=2).value == 125
    assert worksheet.cell(row=4, column=3).value == 2
    assert "Петрова А.И." in worksheet.cell(row=4, column=4).value
    # Автофильтр и закрепление области настроены.
    assert worksheet.freeze_panes == "A4"
    assert worksheet.auto_filter.ref


def test_export_summary_without_project(tmp_path) -> None:
    path = tmp_path / "Сводная2.xlsx"
    ExcelExporter.export_summary([], str(path))
    worksheet = openpyxl.load_workbook(str(path)).active
    assert worksheet.cell(row=1, column=1).value == "ФИО"


def test_export_project_sheets(tmp_path) -> None:
    project = sample_project()
    path = tmp_path / "Полный.xlsx"

    ExcelExporter.export_project(project, str(path))
    workbook = openpyxl.load_workbook(str(path))

    assert workbook.sheetnames[0] == "Сводная"
    assert len(workbook.sheetnames) == 3

    teacher_sheet = workbook["Петрова А.И."]
    values = [row[0] for row in teacher_sheet.iter_rows(values_only=True)]
    assert "Преподаватель: Петрова А.И." in values
    assert "Направление: Математика" in values
    assert "Иванов Иван" in values

    summary_values = [row[0] for row in workbook["Сводная"].iter_rows(values_only=True)]
    assert "Иванов Иван" in summary_values


def test_export_direction(tmp_path) -> None:
    direction = sample_project().teachers[0].directions[0]
    path = tmp_path / "Направление.xlsx"
    ExcelExporter.export_direction(direction, str(path))

    worksheet = openpyxl.load_workbook(str(path)).active
    assert worksheet.title == "Математика"
    assert [cell.value for cell in worksheet[1]] == ["ФИО", "Название конкурса", "Итого"]
    assert worksheet.max_row == 3


def test_export_project_student_subtotals(tmp_path) -> None:
    """На листах преподавателей есть строка «Итого по ученику» с его суммой."""
    project = sample_project()
    path = tmp_path / "Подытоги.xlsx"
    ExcelExporter.export_project(project, str(path))

    worksheet = openpyxl.load_workbook(str(path))["Петрова А.И."]
    rows = list(worksheet.iter_rows(values_only=True))
    subtotals = {row[1]: row[2] for row in rows if row[1] and str(row[1]).startswith("Итого по ученику")}

    assert subtotals == {"Итого по ученику: Иванов Иван": 85, "Итого по ученику: Петрова Анна": 92}
    # Подытог идёт сразу после строк ученика, итог направления — после них.
    labels = [row[1] for row in rows]
    assert labels.index("Итого по ученику: Иванов Иван") < labels.index("Итого по направлению")


def test_export_project_with_invalid_sheet_name(tmp_path) -> None:
    project = sample_project()
    project.teachers[0].name = "Иванов/Пётр: тест"
    path = tmp_path / "Странные имена.xlsx"
    ExcelExporter.export_project(project, str(path))

    workbook = openpyxl.load_workbook(str(path))
    assert len(workbook.sheetnames) == 3
    assert all(":" not in name and "/" not in name for name in workbook.sheetnames)


def test_export_then_import_round_trip(tmp_path) -> None:
    """Экспорт проекта и повторный импорт листа преподавателя сохраняют баллы."""
    project = sample_project()
    path = tmp_path / "Экспорт.xlsx"
    ExcelExporter.export_project(project, str(path))

    # Экспортированный лист преподавателя содержит заголовки «ФИО / Конкурс / Итого»
    # и служебные строки — проверяем, что данные не потерялись в JSON-цикле.
    json_path = tmp_path / "проект.json"
    ProjectManager.save_project(project, str(json_path))
    restored = ProjectManager.load_project(str(json_path))
    assert restored.total_score() == project.total_score()


def test_export_project_score_number_format(tmp_path) -> None:
    """Баллы в экспорте проекта имеют тот же формат «0.##», что и в других экспортах."""
    project = sample_project()
    path = tmp_path / "Формат.xlsx"

    ExcelExporter.export_project(project, str(path))
    workbook = openpyxl.load_workbook(str(path))

    # Сводная: колонка «Итого баллов» (вторая), строки данных со 2-й.
    summary = workbook["Сводная"]
    assert summary.cell(row=2, column=2).value is not None
    assert summary.cell(row=2, column=2).number_format == "0.##"

    # Лист преподавателя: числовые ячейки третьей колонки (баллы и итог
    # направления) отформатированы, текстовые — без числового формата.
    sheet = workbook["Петрова А.И."]
    numeric_formats = [
        cell.number_format
        for row in sheet.iter_rows(min_col=3, max_col=3)
        for cell in row
        if isinstance(cell.value, (int, float))
    ]
    assert numeric_formats, "на листе преподавателя нет числовых ячеек с баллами"
    assert set(numeric_formats) == {"0.##"}
