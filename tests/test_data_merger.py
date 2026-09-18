"""Тесты агрегации данных (этап 4.1 плана)."""

from __future__ import annotations

from evaluation_app.models import Direction, Project
from evaluation_app.services.data_merger import DataMerger
from tests.helpers import sample_project


def test_merge_duplicate_fio() -> None:
    records = DataMerger.merge_by_student(sample_project())
    by_name = {record["student_name"]: record for record in records}

    assert len(records) == 3
    ivanov = by_name["Иванов Иван"]
    assert ivanov["total_score"] == 125
    assert ivanov["entries"] == 2
    assert sorted(ivanov["teachers_involved"]) == ["Петрова А.И.", "Сидоров П.П."]
    assert sorted(ivanov["directions"]) == ["Информатика", "Математика"]
    assert sorted(ivanov["competitions"]) == ["Олимпиада", "Хакатон"]

    assert by_name["Петрова Анна"]["total_score"] == 92
    assert by_name["Кузнецов Лев"]["total_score"] == 70


def test_merge_is_case_and_space_insensitive() -> None:
    project = Project()
    teacher = project.add_teacher(name="Т")
    direction = teacher.add_direction(name="Д")
    direction.add_evaluation(student_name="Иванов  Иван", competition_name="Конкурс", total_score=10)
    direction.add_evaluation(student_name=" иванов иван ", competition_name="Конкурс", total_score=15)

    records = DataMerger.merge_by_student(project)
    assert len(records) == 1
    assert records[0]["total_score"] == 25
    assert records[0]["student_name"] == "Иванов Иван"
    assert records[0]["entries"] == 2


def test_merge_no_duplicates_and_ordering() -> None:
    project = Project()
    teacher = project.add_teacher(name="Т")
    direction = teacher.add_direction(name="Д")
    direction.add_evaluation(student_name="Первый", competition_name="К1", total_score=1)
    direction.add_evaluation(student_name="Второй", competition_name="К2", total_score=2)

    records = DataMerger.merge_by_student(project)
    assert [record["student_name"] for record in records] == ["Первый", "Второй"]


def test_merge_empty_project() -> None:
    assert DataMerger.merge_by_student(Project()) == []
    assert DataMerger.build_statistics(Project())["students"] == 0
    assert DataMerger.format_statistics(DataMerger.build_statistics(Project()))


def test_merge_skips_nameless_records() -> None:
    project = Project()
    teacher = project.add_teacher(name="Т")
    direction = teacher.add_direction(name="Д")
    direction.add_evaluation(student_name="   ", competition_name="К", total_score=50)

    assert DataMerger.merge_by_student(project) == []


def test_merge_direction() -> None:
    direction = Direction(name="Д")
    direction.add_evaluation(student_name="Иванов Иван", competition_name="К1", total_score=10)
    direction.add_evaluation(student_name="Иванов Иван", competition_name="К2", total_score=5)

    records = DataMerger.merge_direction(direction)
    assert len(records) == 1
    assert records[0]["total_score"] == 15
    assert records[0]["competitions"] == ["К1", "К2"]


def test_totals_by_teacher() -> None:
    totals = DataMerger.totals_by_teacher(sample_project())
    assert [item["teacher_name"] for item in totals] == ["Петрова А.И.", "Сидоров П.П."]
    assert totals[0]["total_score"] == 177
    assert totals[0]["directions"] == 1
    assert totals[0]["students"] == 2
    assert totals[1]["entries"] == 2


def test_build_statistics() -> None:
    project = sample_project()
    records = DataMerger.merge_by_student(project)
    stats = DataMerger.build_statistics(project, records)

    assert stats["teachers"] == 2
    assert stats["directions"] == 2
    assert stats["students"] == 3
    assert stats["competitions"] == 2
    assert stats["evaluations"] == 4
    assert stats["total_score"] == 287
    assert stats["max_score"] == 125
    assert stats["min_score"] == 70
    assert stats["average_score"] == round(287 / 3, 2)

    text = DataMerger.format_statistics(stats)
    assert "Преподавателей: 2" in text
    assert "Учеников: 3" in text


def test_statistics_uses_filtered_records() -> None:
    project = sample_project()
    only_ivanov = [record for record in DataMerger.merge_by_student(project) if record["student_name"] == "Иванов Иван"]
    stats = DataMerger.build_statistics(project, only_ivanov)

    assert stats["students"] == 1
    assert stats["total_score"] == 125
    assert stats["evaluations"] == 4  # статистика по проекту целиком


def test_teacher_records() -> None:
    project = sample_project()
    teacher = project.teachers[0]
    records = DataMerger.teacher_records(teacher)
    assert {record["student_name"] for record in records} == {"Иванов Иван", "Петрова Анна"}
    assert teacher.total_score() == 177


def test_teacher_without_directions() -> None:
    project = Project()
    project.add_teacher(name="Пустой")
    assert DataMerger.merge_by_student(project) == []
    assert DataMerger.totals_by_teacher(project)[0]["total_score"] == 0
