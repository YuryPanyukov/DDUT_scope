"""Объединение (агрегация) данных учеников."""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from evaluation_app.models import Direction, Project, Teacher
from evaluation_app.utils.validators import name_key, normalize_name


def _append_unique(target: list, value: Any) -> None:
    """Добавляет значение в список, если его там ещё нет."""
    if value is None:
        return
    if isinstance(value, str):
        value = value.strip()
    if not value:
        return
    if value not in target:
        target.append(value)


class DataMerger:
    """Агрегация оценок по ФИО ученика (см. п. 4.1 плана)."""

    @staticmethod
    def empty_record(student_name: str) -> dict:
        """Пустая запись агрегированной таблицы."""
        return {
            "student_name": student_name,
            "total_score": 0.0,
            "entries": 0,
            "teachers_involved": [],
            "directions": [],
            "competitions": [],
        }

    @staticmethod
    def merge_by_student(project: Project) -> list[dict]:
        """Формирует сводную таблицу: ФИО ученика → сумма баллов.

        Одинаковые ФИО (в том числе отличающиеся регистром или лишними
        пробелами) суммируются по всем преподавателям и направлениям.
        """
        groups: dict[str, dict] = {}

        for teacher in project.teachers:
            for direction in teacher.directions:
                for evaluation in direction.evaluations:
                    display_name = normalize_name(evaluation.student_name)
                    if not display_name:
                        continue

                    key = name_key(display_name)
                    record = groups.get(key)
                    if record is None:
                        record = DataMerger.empty_record(display_name)
                        groups[key] = record

                    record["total_score"] = float(record["total_score"]) + float(
                        evaluation.total_score or 0.0
                    )
                    record["entries"] = int(record["entries"]) + 1
                    _append_unique(record["teachers_involved"], teacher.name)
                    _append_unique(record["directions"], direction.name)
                    _append_unique(record["competitions"], evaluation.competition_name)

        return list(groups.values())

    @staticmethod
    def merge_direction(direction: Direction) -> list[dict]:
        """Сводная таблица (ФИО → сумма баллов) внутри одного направления.

        Используется только в тестах: интерфейс показывает сводную по всему
        проекту (:meth:`merge_by_student`).
        """
        group: dict[str, dict] = {}
        for evaluation in direction.evaluations:
            display_name = normalize_name(evaluation.student_name)
            if not display_name:
                continue
            key = name_key(display_name)
            record = group.get(key)
            if record is None:
                record = DataMerger.empty_record(display_name)
                group[key] = record
            record["total_score"] = float(record["total_score"]) + float(
                evaluation.total_score or 0.0
            )
            record["entries"] = int(record["entries"]) + 1
            _append_unique(record["directions"], direction.name)
            _append_unique(record["competitions"], evaluation.competition_name)
        return list(group.values())

    @staticmethod
    def teacher_records(teacher: Teacher) -> list[dict]:
        """Сводные данные по одному преподавателю (все его направления).

        Используется только в тестах — обёртка над :meth:`merge_by_student`
        для проекта из одного преподавателя.
        """
        records = DataMerger.merge_by_student(Project(teachers=[teacher]))
        return records

    @staticmethod
    def totals_by_teacher(project: Project) -> list[dict]:
        """Баллы и количество оценок в разрезе преподавателей.

        Используется только в тестах: в интерфейсе такие итоги выводятся
        в списке преподавателей (:class:`TeacherListWidget`).
        """
        result = []
        for teacher in project.teachers:
            result.append(
                {
                    "teacher_name": teacher.name,
                    "directions": len(teacher.directions),
                    "entries": len(teacher.evaluations()),
                    "students": len(teacher.student_names()),
                    "total_score": teacher.total_score(),
                }
            )
        return result

    @staticmethod
    def build_statistics(
        project: Project, records: Iterable[dict] | None = None
    ) -> dict:
        """Статистика по проекту (см. п. 8.6 плана)."""
        all_records = list(records) if records is not None else DataMerger.merge_by_student(project)
        scores = [float(record.get("total_score") or 0.0) for record in all_records]

        competitions: list[str] = []
        directions: list[str] = []
        evaluations_count = 0
        for teacher in project.teachers:
            for direction in teacher.directions:
                _append_unique(directions, direction.name)
                for evaluation in direction.evaluations:
                    evaluations_count += 1
                    _append_unique(competitions, evaluation.competition_name)

        return {
            "project_name": project.project_name,
            "academic_year": project.academic_year,
            "teachers": len(project.teachers),
            "directions": len(directions),
            "students": len(all_records),
            "competitions": len(competitions),
            "evaluations": evaluations_count,
            "total_score": float(sum(scores)),
            "average_score": round(sum(scores) / len(scores), 2) if scores else 0.0,
            "min_score": min(scores) if scores else 0.0,
            "max_score": max(scores) if scores else 0.0,
        }

    @staticmethod
    def format_statistics(stats: dict) -> str:
        """Строка статистики для строки состояния."""
        if not stats:
            return ""
        return (
            f"Преподавателей: {stats['teachers']} | "
            f"Направлений: {stats['directions']} | "
            f"Учеников: {stats['students']} | "
            f"Конкурсов: {stats['competitions']} | "
            f"Оценок: {stats['evaluations']} | "
            f"Сумма баллов: {stats['total_score']:g} | "
            f"Средний балл: {stats['average_score']:g} | "
            f"Мин: {stats['min_score']:g} | Макс: {stats['max_score']:g}"
        )
