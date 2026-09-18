"""Тесты сортировки (этап 4.2 плана)."""

from __future__ import annotations

from evaluation_app.services.data_merger import DataMerger
from evaluation_app.services.data_sorter import SORT_KEYS, DataSorter
from tests.helpers import sample_project


def records() -> list[dict]:
    return DataMerger.merge_by_student(sample_project())


def test_sort_by_name() -> None:
    names = [record["student_name"] for record in DataSorter.sort_by_name(records())]
    assert names == ["Иванов Иван", "Кузнецов Лев", "Петрова Анна"]

    reverse_names = [record["student_name"] for record in DataSorter.sort_by_name(records(), reverse=True)]
    assert reverse_names == list(reversed(names))


def test_sort_by_name_ignores_case() -> None:
    data = [
        {"student_name": "яблоков Пётр", "total_score": 1},
        {"student_name": "Иванов Иван", "total_score": 1},
    ]
    assert [r["student_name"] for r in DataSorter.sort_by_name(data)] == ["Иванов Иван", "яблоков Пётр"]


def test_sort_by_score() -> None:
    scores = [record["total_score"] for record in DataSorter.sort_by_score(records())]
    assert scores == [125, 92, 70]
    assert [record["total_score"] for record in DataSorter.sort_by_score(records(), reverse=False)] == [70, 92, 125]


def test_sort_by_competitions_and_entries() -> None:
    by_competitions = DataSorter.sort_by_competitions(records())
    assert by_competitions[0]["student_name"] == "Иванов Иван"
    assert len(by_competitions[0]["competitions"]) == 2

    by_entries = DataSorter.sort_by_entries(records())
    assert by_entries[0]["entries"] == 2


def test_universal_sort_and_defaults() -> None:
    assert [r["student_name"] for r in DataSorter.sort(records(), key="name", reverse=False)] == [
        "Иванов Иван",
        "Кузнецов Лев",
        "Петрова Анна",
    ]
    # Неизвестный ключ — сортировка по баллам по убыванию.
    assert DataSorter.sort(records(), key="нет-ключа")[0]["total_score"] == 125
    assert DataSorter.default_reverse("name") is False
    assert DataSorter.default_reverse("score") is True


def test_sort_does_not_mutate_source() -> None:
    source = records()
    original = [record["student_name"] for record in source]
    DataSorter.sort_by_score(source)
    assert [record["student_name"] for record in source] == original


def test_sort_keys_are_available_for_ui() -> None:
    assert set(SORT_KEYS) == {"name", "score", "competitions", "entries"}


def test_sort_handles_missing_fields() -> None:
    data = [{"student_name": "А"}, {"student_name": "Б", "total_score": None, "competitions": None}]
    assert len(DataSorter.sort_by_score(data)) == 2
    assert len(DataSorter.sort_by_name(data)) == 2
