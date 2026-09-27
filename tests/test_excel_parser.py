"""Тесты парсинга Excel (этап 3 плана)."""

from __future__ import annotations

import openpyxl
import pytest

from evaluation_app.services.excel_parser import ExcelParseError, ExcelParser
from tests.helpers import make_workbook

STANDARD_ROWS = [
    ["ФИО", "Название конкурса", "Итого"],
    ["Иванов Иван", "Олимпиада по математике", 85],
    ["Петрова Анна", "Олимпиада по математике", 92],
]


def test_parse_single_sheet(tmp_path) -> None:
    path = make_workbook(tmp_path / "p.xlsx", {"Математика": STANDARD_ROWS})
    parser = ExcelParser(path)
    parser.load()

    assert parser.sheet_names == ["Математика"]
    result = parser.parse_sheet("Математика")
    assert result.header_row == 1
    assert result.evaluations_count == 2
    assert result.total_score == 177
    assert result.direction.evaluations[0].student_name == "Иванов Иван"
    assert result.columns["total_score"] == 3


def test_parse_all_teachers_groups_sheets(tmp_path) -> None:
    path = make_workbook(
        tmp_path / "Петрова_2024.xlsx",
        {
            "Математика": STANDARD_ROWS,
            "Информатика": [
                ["Ученик", "Конкурс", "Баллы"],
                ["Кузнецов Лев", "Хакатон", 70],
            ],
        },
    )
    parser = ExcelParser(path, teacher_name="Петрова А.И.")
    teachers = parser.parse_all_teachers()

    assert len(teachers) == 1
    assert teachers[0].name == "Петрова А.И."
    assert [d.name for d in teachers[0].directions] == ["Математика", "Информатика"]
    assert teachers[0].total_score() == 247


def test_parse_all_teachers_default_name_from_file(tmp_path) -> None:
    path = make_workbook(tmp_path / "Петрова_2024.xlsx", {"Математика": STANDARD_ROWS})
    teachers = ExcelParser(path).parse_all_teachers()
    assert teachers[0].name == "Петрова"


def test_parse_all_teachers_with_mapping_and_skip(tmp_path) -> None:
    path = make_workbook(
        tmp_path / "file.xlsx",
        {
            "Лист1": STANDARD_ROWS,
            "Лист2": [["ФИО", "Конкурс", "Итого"], ["Сидоров Пётр", "Конкурс", 10]],
            "Лист3": [["ФИО", "Конкурс", "Итого"], ["Игнор Игнор", "Конкурс", 10]],
        },
    )
    parser = ExcelParser(path)
    teachers = parser.parse_all_teachers(
        {"Лист1": "Петрова А.И.", "Лист2": "Сидоров П.П.", "Лист3": None}
    )

    assert sorted(t.name for t in teachers) == ["Петрова А.И.", "Сидоров П.П."]
    assert all("Лист3" not in [d.name for d in t.directions] for t in teachers)


def test_headers_can_be_shuffled_or_named_differently(tmp_path) -> None:
    path = make_workbook(
        tmp_path / "aliases.xlsx",
        {
            "Лист": [
                ["Итого баллов", "Комментарий", "Фамилия Имя Отчество", "Мероприятие"],
                [80, "ок", "Иванов Иван", "Олимпиада"],
            ]
        },
    )
    result = ExcelParser(path).parse_sheet("Лист")

    assert result.columns == {
        "total_score": 1,
        "student_name": 3,
        "competition_name": 4,
    }
    evaluation = result.direction.evaluations[0]
    assert evaluation.competition_name == "Олимпиада"
    assert evaluation.total_score == 80


def test_header_row_is_detected_after_title_rows(tmp_path) -> None:
    path = make_workbook(
        tmp_path / "title.xlsx",
        {
            "Лист": [
                ["Отчёт преподавателя за 2025 год"],
                [],
                ["ФИО", "Название конкурса", "Итого"],
                ["Иванов Иван", "Олимпиада", 55],
            ]
        },
    )
    result = ExcelParser(path).parse_sheet("Лист")
    assert result.header_row == 3
    assert result.evaluations_count == 1


def test_problem_rows_are_skipped_with_warnings(tmp_path) -> None:
    path = make_workbook(
        tmp_path / "bad.xlsx",
        {
            "Лист": [
                ["ФИО", "Название конкурса", "Итого"],
                ["Иванов Иван", "Олимпиада", 85],
                [None, None, None],
                ["", "", ""],
                [None, "Олимпиада", 50],
                ["Петрова Анна", "Олимпиада", "нет данных"],
                ["Кузнецов Лев", "", 40],
                ["Сидоров Пётр", "Олимпиада", "1 000"],
            ]
        },
    )
    result = ExcelParser(path).parse_sheet("Лист")

    assert result.evaluations_count == 3
    assert result.skipped_rows == [5, 6]
    assert any("не является числом" in warning for warning in result.warnings)
    assert any("не указано ФИО" in warning for warning in result.warnings)
    assert any("не указано название конкурса" in warning for warning in result.warnings)
    # Пустые строки не считаются проблемными.
    assert 3 not in result.skipped_rows and 4 not in result.skipped_rows
    assert result.direction.evaluations[-1].total_score == 1000


def test_numeric_and_comma_scores(tmp_path) -> None:
    path = make_workbook(
        tmp_path / "scores.xlsx",
        {
            "Лист": [
                ["ФИО", "Конкурс", "Итого"],
                ["Первый Ученик", "Конкурс", "85,5"],
                ["Второй Ученик", "Конкурс", "12 баллов"],
                ["Третий Ученик", "Конкурс", 40.25],
            ]
        },
    )
    result = ExcelParser(path).parse_sheet("Лист")
    assert [e.total_score for e in result.direction.evaluations] == [85.5, 12.0, 40.25]


def test_rows_mode_keeps_student_with_empty_score(tmp_path) -> None:
    """Пустая ячейка баллов — 0 баллов: ученик не теряется из-за нуля (задача 2)."""
    path = make_workbook(
        tmp_path / "zero.xlsx",
        {
            "Лист": [
                ["ФИО", "Название конкурса", "Итого"],
                ["Иванов Иван", "Олимпиада", 85],
                ["Петров Пётр", "Олимпиада", None],
                ["Сидоров Сидор", "Олимпиада", ""],
            ]
        },
    )
    result = ExcelParser(path).parse_sheet("Лист")

    scores = {e.student_name: e.total_score for e in result.direction.evaluations}
    assert scores == {"Иванов Иван": 85.0, "Петров Пётр": 0.0, "Сидоров Сидор": 0.0}
    assert result.rows_parsed == 3
    assert not any("не является числом" in warning for warning in result.warnings)


def test_empty_sheet_produces_warning(tmp_path) -> None:
    path = make_workbook(tmp_path / "empty.xlsx", {"Лист": [["ФИО", "Конкурс", "Итого"]]})
    result = ExcelParser(path).parse_sheet("Лист")
    assert result.evaluations_count == 0
    assert any("ни одной корректной строки" in warning for warning in result.warnings)


def test_missing_headers_raise_error(tmp_path) -> None:
    path = make_workbook(
        tmp_path / "unknown.xlsx",
        {"Лист": [["Столбец А", "Столбец Б"], ["значение", "значение"]]},
    )
    with pytest.raises(ExcelParseError):
        ExcelParser(path).parse_sheet("Лист")


def test_manual_overrides(tmp_path) -> None:
    # Столбцы без узнаваемых заголовков: указываем структуру вручную.
    path = make_workbook(
        tmp_path / "manual2.xlsx",
        {
            "Лист": [
                ["Отчёт"],
                ["Кто", "Конкурс", "Баллы"],
                ["Иванов Иван", "Олимпиада", 85],
            ]
        },
    )
    parser = ExcelParser(path)
    parser.set_overrides(header_row=2, columns={"student_name": 1, "competition_name": 2, "total_score": 3})
    result = parser.parse_sheet("Лист")
    assert result.evaluations_count == 1
    assert result.columns["student_name"] == 1


def test_detect_header_with_fixed_row_error(tmp_path) -> None:
    path = make_workbook(tmp_path / "fixed.xlsx", {"Лист": [["A", "B"], ["1", "2"]]})
    parser = ExcelParser(path)
    with pytest.raises(ExcelParseError):
        parser.detect_header(parser.worksheet("Лист"), header_row=1)


def test_file_not_found_and_bad_extension(tmp_path) -> None:
    with pytest.raises(ExcelParseError):
        ExcelParser(tmp_path / "нет-файла.xlsx").load()

    text_file = tmp_path / "данные.txt"
    text_file.write_text("не excel", encoding="utf-8")
    with pytest.raises(ExcelParseError):
        ExcelParser(str(text_file)).load()


def test_unknown_sheet_raises(tmp_path) -> None:
    path = make_workbook(tmp_path / "one.xlsx", {"Математика": STANDARD_ROWS})
    parser = ExcelParser(path)
    with pytest.raises(ExcelParseError):
        parser.worksheet("НетТакого")


def test_preview_rows(tmp_path) -> None:
    path = make_workbook(tmp_path / "preview.xlsx", {"Математика": STANDARD_ROWS})
    parser = ExcelParser(path)
    rows = parser.preview_rows("Математика", limit=2)
    assert rows == [["ФИО", "Название конкурса", "Итого"], ["Иванов Иван", "Олимпиада по математике", 85]]


def test_parse_all_collects_errors_without_strict(tmp_path) -> None:
    path = make_workbook(
        tmp_path / "mixed.xlsx",
        {"Хороший": STANDARD_ROWS, "Плохой": [["Нет", "Нужных", "Столбцов"]]},
    )
    parser = ExcelParser(path)
    results = parser.parse_all()

    assert set(results) == {"Хороший", "Плохой"}
    assert results["Хороший"].evaluations_count == 2
    assert results["Плохой"].evaluations_count == 0
    assert parser.warnings


def test_formulas_are_read_as_values(tmp_path) -> None:
    path = tmp_path / "formula.xlsx"
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Лист"
    worksheet.append(["ФИО", "Конкурс", "Итого"])
    worksheet.append(["Иванов Иван", "Олимпиада", "=40+45"])
    workbook.save(str(path))

    # Без пересчёта формул Excel значение недоступно: строка не выбрасывается
    # (баллы учтены как 0), о пустой ячейке есть предупреждение, парсер не падает.
    result = ExcelParser(path).parse_sheet("Лист")
    assert result.evaluations_count == 1
    assert result.direction.evaluations[0].total_score == 0
    assert any("не указано значение баллов" in warning for warning in result.warnings)
