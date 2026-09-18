"""Тесты сохранения/загрузки проекта (этап 4.3 плана)."""

from __future__ import annotations

import json

import pytest

from evaluation_app.models import Project, Teacher
from evaluation_app.services.project_manager import ProjectError, ProjectManager
from tests.helpers import sample_project


def test_save_and_load_cycle(tmp_path) -> None:
    project = sample_project()
    path = tmp_path / "проект.json"

    ProjectManager.save_project(project, str(path))
    assert path.is_file()

    data = json.loads(path.read_text(encoding="utf-8"))
    assert data["project_name"] == "Тестовый проект"
    assert data["format_version"] == 1
    assert "saved_at" in data
    assert len(data["teachers"]) == 2

    restored = ProjectManager.load_project(str(path))
    assert restored.to_dict() == project.to_dict()
    assert restored.teachers[0].directions[0].evaluations[0].total_score == 85


def test_save_creates_directories(tmp_path) -> None:
    path = tmp_path / "вложенный" / "каталог" / "проект.json"
    ProjectManager.save_project(sample_project(), str(path))
    assert path.is_file()


def test_load_missing_file_raises(tmp_path) -> None:
    with pytest.raises(ProjectError) as exc:
        ProjectManager.load_project(str(tmp_path / "нет.json"))
    assert "не найден" in str(exc.value)


def test_load_broken_json_raises(tmp_path) -> None:
    path = tmp_path / "broken.json"
    path.write_text("{ это не json", encoding="utf-8")
    with pytest.raises(ProjectError):
        ProjectManager.load_project(str(path))


def test_load_non_dict_json_raises(tmp_path) -> None:
    path = tmp_path / "list.json"
    path.write_text("[1, 2, 3]", encoding="utf-8")
    with pytest.raises(ProjectError):
        ProjectManager.load_project(str(path))


def test_load_tolerates_partial_data(tmp_path) -> None:
    path = tmp_path / "partial.json"
    path.write_text(
        json.dumps({"project_name": "Частичный", "teachers": [{"name": "Т", "directions": [{"name": "Д"}]}]}),
        encoding="utf-8",
    )
    project = ProjectManager.load_project(str(path))
    assert project.project_name == "Частичный"
    assert project.teachers[0].directions[0].evaluations == []


def test_save_without_path_raises() -> None:
    with pytest.raises(ProjectError):
        ProjectManager.save_project(sample_project(), "")


def test_update_teacher_replaces_all_data() -> None:
    project = sample_project()
    teacher = project.teachers[0]
    original_id = teacher.id

    new_teacher = Teacher(name="Петрова А.И.")
    direction = new_teacher.add_direction(name="Математика")
    direction.add_evaluation(student_name="Новый Ученик", competition_name="Конкурс", total_score=99)

    assert ProjectManager.update_teacher(project, original_id, new_teacher)
    updated = project.get_teacher(original_id)
    assert updated.id == original_id
    assert [e.student_name for e in updated.evaluations()] == ["Новый Ученик"]
    assert updated.total_score() == 99
    assert not ProjectManager.update_teacher(project, "нет-id", new_teacher)


def test_replace_teacher_directions() -> None:
    project = sample_project()
    teacher = project.teachers[0]

    new_direction = Teacher(name="X").add_direction(name="Физика")
    assert ProjectManager.replace_teacher_directions(project, teacher.id, [new_direction])
    assert [d.name for d in project.get_teacher(teacher.id).directions] == ["Физика"]
    assert not ProjectManager.replace_teacher_directions(project, "нет-id", [])


def test_remove_teacher() -> None:
    project = sample_project()
    teacher_id = project.teachers[0].id

    assert ProjectManager.remove_teacher(project, teacher_id)
    assert project.get_teacher(teacher_id) is None
    assert not ProjectManager.remove_teacher(project, teacher_id)


def test_find_teacher_by_name() -> None:
    project = sample_project()
    assert ProjectManager.find_teacher(project, " петрова  а.и. ") is project.teachers[0]
    assert ProjectManager.find_teacher(project, "Нет Такого") is None


def test_round_trip_in_memory_project() -> None:
    project = Project(project_name="Пустой", academic_year="2025-2026")
    clone = Project.from_dict(project.to_dict())
    assert clone.is_empty()
    assert clone.project_name == "Пустой"
