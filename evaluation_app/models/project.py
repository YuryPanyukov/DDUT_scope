"""Модель проекта — контейнер всех данных приложения."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, field

from evaluation_app.models.teacher import Teacher
from evaluation_app.utils.validators import name_key


@dataclass
class Project:
    """Проект: список преподавателей с направлениями и оценками."""

    project_name: str = ""
    academic_year: str = ""
    teachers: list[Teacher] = field(default_factory=list)

    # --- Преподаватели -------------------------------------------------------

    def add_teacher(self, teacher: Teacher | None = None, *, name: str = "") -> Teacher:
        """Добавляет преподавателя и возвращает его."""
        if teacher is None:
            teacher = Teacher(name=name)
        self.teachers.append(teacher)
        return teacher

    def remove_teacher(self, teacher_id: str) -> bool:
        """Удаляет преподавателя по идентификатору."""
        for index, teacher in enumerate(self.teachers):
            if teacher.id == teacher_id:
                del self.teachers[index]
                return True
        return False

    def get_teacher(self, teacher_id: str) -> Teacher | None:
        """Возвращает преподавателя по идентификатору."""
        for teacher in self.teachers:
            if teacher.id == teacher_id:
                return teacher
        return None

    def get_teacher_index(self, teacher_id: str) -> int:
        """Индекс преподавателя в списке или -1."""
        for index, teacher in enumerate(self.teachers):
            if teacher.id == teacher_id:
                return index
        return -1

    def find_teacher(self, name: str) -> Teacher | None:
        """Находит преподавателя по ФИО (без учёта регистра)."""
        key = name_key(name)
        if not key:
            return None
        for teacher in self.teachers:
            if name_key(teacher.name) == key:
                return teacher
        return None

    def get_or_create_teacher(self, name: str) -> Teacher:
        """Возвращает преподавателя с таким ФИО, создавая его при необходимости.

        Используется только в тестах: импорт в интерфейсе сам ищет
        преподавателя (:meth:`find_teacher`) и добавляет нового через
        :meth:`add_teacher`.
        """
        teacher = self.find_teacher(name)
        if teacher is None:
            teacher = self.add_teacher(name=name)
        return teacher

    def replace_teacher(self, teacher_id: str, new_teacher: Teacher) -> bool:
        """Заменяет преподавателя целиком (старые данные удаляются).

        Используется только в тестах — через
        :meth:`ProjectManager.update_teacher`. Обновление данных из Excel
        меняет только направления (:meth:`replace_teacher_directions`).
        """
        index = self.get_teacher_index(teacher_id)
        if index < 0:
            return False
        if not new_teacher.id or new_teacher.id != teacher_id:
            new_teacher.id = teacher_id
        self.teachers[index] = new_teacher
        return True

    def replace_teacher_directions(self, teacher_id: str, directions: Iterable) -> bool:
        """Заменяет только направления преподавателя."""
        teacher = self.get_teacher(teacher_id)
        if teacher is None:
            return False
        teacher.replace_directions(directions)
        return True

    def rename_teacher(self, teacher_id: str, new_name: str) -> bool:
        """Переименовывает преподавателя."""
        teacher = self.get_teacher(teacher_id)
        if teacher is None or not new_name.strip():
            return False
        new_name = new_name.strip()
        teacher.name = new_name
        return True

    # --- Агрегаты ------------------------------------------------------------

    def directions(self) -> list:
        """Все направления проекта (по всем преподавателям).

        Используется только в тестах: интерфейс и агрегаты обходят
        ``teacher.directions`` напрямую.
        """
        return [
            direction
            for teacher in self.teachers
            for direction in teacher.directions
        ]

    def evaluations(self) -> list:
        """Все оценки проекта."""
        return [
            evaluation
            for teacher in self.teachers
            for evaluation in teacher.evaluations()
        ]

    def student_names(self) -> list[str]:
        """Уникальные ФИО учеников проекта в порядке появления.

        Используется только в тестах: подписи в интерфейсе считаются на уровне
        преподавателя, а сводная по проекту — в :meth:`DataMerger.merge_by_student`.
        """
        result: list[str] = []
        for evaluation in self.evaluations():
            if evaluation.student_name and evaluation.student_name not in result:
                result.append(evaluation.student_name)
        return result

    def total_score(self) -> float:
        """Сумма всех баллов проекта.

        Используется только в тестах: статистика и экспорт считают сумму сами,
        чтобы не обходить проект лишний раз.
        """
        return float(sum(evaluation.total_score for evaluation in self.evaluations()))

    def is_empty(self) -> bool:
        """Есть ли в проекте хоть какие-то данные."""
        return not self.teachers

    # --- Сериализация --------------------------------------------------------

    def to_dict(self) -> dict:
        """Сериализация проекта в словарь."""
        return {
            "project_name": self.project_name,
            "academic_year": self.academic_year,
            "teachers": [teacher.to_dict() for teacher in self.teachers],
        }

    @classmethod
    def from_dict(cls, data: dict) -> Project:
        """Восстановление проекта из словаря."""
        data = data or {}
        project = cls(
            project_name=str(data.get("project_name") or ""),
            academic_year=str(data.get("academic_year") or ""),
        )
        for item in data.get("teachers") or []:
            if isinstance(item, dict):
                project.teachers.append(Teacher.from_dict(item))
        return project

    def clone(self) -> Project:
        """Глубокая копия проекта, независимая от оригинала.

        Используется только в тестах: история изменений хранит снимки через
        :meth:`to_dict`/:meth:`from_dict`.
        """
        return Project.from_dict(self.to_dict())
