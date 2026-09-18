"""Тесты валидации и разбора значений (этап 7 плана)."""

from __future__ import annotations

import openpyxl
import pytest

from evaluation_app.utils.validators import (
    DataValidator,
    detect_columns,
    match_column,
    missing_fields,
    name_key,
    normalize_header,
    normalize_name,
    parse_score,
)


@pytest.mark.parametrize(
    "header,expected",
    [
        ("ФИО", "student_name"),
        ("Ученик", "student_name"),
        ("Фамилия Имя Отчество", "student_name"),
        ("обучающийся", "student_name"),
        ("Название конкурса", "competition_name"),
        ("Конкурс", "competition_name"),
        ("Итого", "total_score"),
        ("Итого баллов", "total_score"),
        ("Баллы", "total_score"),
        ("Количество баллов ", "total_score"),
        ("Комментарий", None),
        ("", None),
        (None, None),
    ],
)
def test_match_column(header, expected) -> None:
    assert match_column(header) == expected


def test_detect_columns_any_order() -> None:
    mapping = detect_columns(["Итого", "Пояснение", "Ученик", "Конкурс"])
    assert mapping == {"total_score": 1, "student_name": 3, "competition_name": 4}
    assert missing_fields(mapping) == []


def test_missing_fields() -> None:
    assert missing_fields({"student_name": 1}) == ["total_score"]


def test_normalize_helpers() -> None:
    assert normalize_header("  Итого\n баллов ") == "итого баллов"
    assert normalize_header("ИТОГО БАЛЛОВ") == "итого баллов"
    assert normalize_name("  Иванов   Иван ") == "Иванов Иван"
    assert name_key("Иванов Иван") == name_key(" иванов  иван ")


@pytest.mark.parametrize(
    "value,expected",
    [
        (85, 85.0),
        ("85", 85.0),
        ("85,5", 85.5),
        ("85.5", 85.5),
        ("12 баллов", 12.0),
        ("  100  ", 100.0),
        ("1 234", 1234.0),
        ("", None),
        ("нет данных", None),
        ("-", None),
        (None, None),
        (True, None),
    ],
)
def test_parse_score(value, expected) -> None:
    assert parse_score(value) == expected


def test_validate_score() -> None:
    assert DataValidator.validate_score("90")
    assert not DataValidator.validate_score("победитель")


def test_validate_excel_headers_with_worksheet() -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.append(["ФИО", "Название конкурса", "Итого"])
    worksheet.append(["Иванов Иван", "Олимпиада", 85])
    assert DataValidator.validate_excel_headers(worksheet)

    worksheet_bad = workbook.create_sheet("Плохой")
    worksheet_bad.append(["ФИО", "Комментарий"])
    assert not DataValidator.validate_excel_headers(worksheet_bad)
    assert "Итого баллов" in DataValidator.describe_missing(detect_columns(["ФИО", "Комментарий"]))
