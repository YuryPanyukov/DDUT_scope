"""Тесты моделей данных (этап 2 плана)."""

from __future__ import annotations

from evaluation_app.models import Direction, Evaluation, Project, Teacher


def test_evaluation_normalizes_values() -> None:
    evaluation = Evaluation(student_name="  Иванов   Иван ", competition_name=" Олимпиада ", total_score="85,5")
    assert evaluation.student_name == "Иванов Иван"
    assert evaluation.competition_name == "Олимпиада"
    assert evaluation.total_score == 85.5
    assert evaluation.id


def test_evaluation_invalid_score_becomes_zero() -> None:
    assert Evaluation(student_name="Тест", total_score="нет данных").total_score == 0.0


def test_direction_add_and_remove() -> None:
    direction = Direction(name="Математика")
    added = direction.add_evaluation(student_name="Иванов Иван", competition_name="Олимпиада", total_score=85)
    direction.add_evaluation(student_name="Петрова Анна", competition_name="Олимпиада", total_score=92)

    assert len(direction) == 2
    assert direction.total_score() == 177
    assert direction.student_names() == ["Иванов Иван", "Петрова Анна"]
    assert direction.get_evaluation(added.id) is added
    assert direction.remove_evaluation(added.id)
    assert not direction.remove_evaluation("нет-такого-id")
    assert len(direction) == 1


def test_teacher_directions() -> None:
    teacher = Teacher(name="Петрова А.И.")
    teacher.add_direction(name="Математика")
    math = teacher.get_direction("математика")
    assert math is not None

    math.add_evaluation(student_name="Иванов Иван", competition_name="Олимпиада", total_score=85)
    # Повторное добавление с тем же именем заменяет содержимое.
    teacher.add_direction(Direction(name="Математика"))
    assert len(teacher.directions) == 1

    physics = teacher.add_direction(name="Физика")
    physics.add_evaluation(student_name="Петрова Анна", competition_name="Конкурс", total_score=60)

    assert len(teacher.directions) == 2
    assert teacher.total_score() == 60
    assert teacher.student_names() == ["Петрова Анна"]
    assert teacher.matches_name("  петрова  а.и. ")
    assert teacher.remove_direction(physics.id)
    assert len(teacher.directions) == 1


def test_project_teacher_management() -> None:
    project = Project(project_name="П", academic_year="2025-2026")
    teacher = project.add_teacher(name="Иванов И.И.")

    assert project.get_teacher(teacher.id) is teacher
    assert project.find_teacher("иванов и.и.") is teacher
    assert project.get_or_create_teacher("Иванов И.И.") is teacher

    other = project.get_or_create_teacher("Сидоров П.П.")
    assert other is not teacher
    assert len(project.teachers) == 2

    project.rename_teacher(teacher.id, "Иванов Пётр Иванович")
    assert project.get_teacher(teacher.id).name == "Иванов Пётр Иванович"
    assert not project.rename_teacher(teacher.id, "   ")

    replacement = Teacher(name="Новый")
    assert project.replace_teacher(teacher.id, replacement)
    assert project.get_teacher(teacher.id).name == "Новый"
    assert replacement.id == teacher.id
    assert not project.replace_teacher("нет-id", Teacher(name="X"))

    assert project.remove_teacher(other.id)
    assert not project.remove_teacher(other.id)


def test_project_aggregates_and_clone() -> None:
    project = Project(project_name="П")
    teacher = project.add_teacher(name="Петрова А.И.")
    direction = teacher.add_direction(name="Математика")
    direction.add_evaluation(student_name="Иванов Иван", competition_name="Олимпиада", total_score=85)
    direction.add_evaluation(student_name="Петрова Анна", competition_name="Олимпиада", total_score=15)

    assert project.total_score() == 100
    assert project.student_names() == ["Иванов Иван", "Петрова Анна"]
    assert len(project.directions()) == 1
    assert len(project.evaluations()) == 2
    assert not project.is_empty()

    clone = project.clone()
    clone.teachers[0].name = "Изменено"
    clone.evaluations()[0].total_score = 999
    assert project.teachers[0].name == "Петрова А.И."
    assert project.evaluations()[0].total_score == 85


def test_project_serialization_round_trip() -> None:
    project = Project(project_name="Тест", academic_year="2025-2026")
    teacher = project.add_teacher(name="Петрова А.И.")
    direction = teacher.add_direction(name="Математика")
    direction.add_evaluation(student_name="Иванов Иван", competition_name="Олимпиада", total_score=85.5)

    data = project.to_dict()
    restored = Project.from_dict(data)

    assert restored.to_dict() == data
    assert restored.teachers[0].directions[0].evaluations[0].total_score == 85.5
    assert restored.teachers[0].directions[0].id == direction.id

    # Устойчивость к неполным данным.
    assert Project.from_dict({}).teachers == []
    assert Evaluation.from_dict({"student_name": "X"}).total_score == 0.0
    direction_data = {"name": "D", "evaluations": [{"student_name": "X", "total_score": "10"}]}
    assert Direction.from_dict(direction_data).total_score() == 10
