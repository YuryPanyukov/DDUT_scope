"""Вспомогательные функции для тестов."""

from __future__ import annotations

from collections.abc import Sequence

import openpyxl

from evaluation_app.models import Direction, Evaluation, Project, Teacher


def make_workbook(file_path: str, sheets: dict[str, Sequence[Sequence]]) -> str:
    """Создаёт Excel-файл: словарь «имя листа → строки»."""
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    for sheet_name, rows in sheets.items():
        worksheet = workbook.create_sheet(title=sheet_name)
        for row in rows:
            worksheet.append(list(row))
    workbook.save(str(file_path))
    return str(file_path)


def sample_project() -> Project:
    """Небольшой проект с дублирующимися ФИО и разными преподавателями."""
    teacher_one = Teacher(name="Петрова А.И.")
    direction_one = teacher_one.add_direction(name="Математика")
    direction_one.add_evaluation(student_name="Иванов Иван", competition_name="Олимпиада", total_score=85)
    direction_one.add_evaluation(student_name="Петрова Анна", competition_name="Олимпиада", total_score=92)

    teacher_two = Teacher(name="Сидоров П.П.")
    direction_two = teacher_two.add_direction(name="Информатика")
    direction_two.add_evaluation(student_name="иванов иван", competition_name="Хакатон", total_score=40)
    direction_two.add_evaluation(student_name="Кузнецов Лев", competition_name="Хакатон", total_score=70)

    project = Project(project_name="Тестовый проект", academic_year="2025-2026")
    project.teachers.extend([teacher_one, teacher_two])
    return project


def direction(rows: Sequence[tuple]) -> Direction:
    """Создаёт направление из троек (ФИО, конкурс, баллы)."""
    result = Direction(name="Тест")
    for student, competition, score in rows:
        result.evaluations.append(
            Evaluation(student_name=student, competition_name=competition, total_score=score)
        )
    return result
