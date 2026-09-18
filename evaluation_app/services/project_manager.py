"""Сохранение и загрузка проекта (см. п. 4.3 плана)."""

from __future__ import annotations

import os
from collections.abc import Iterable
from datetime import datetime

from evaluation_app.models import Project, Teacher
from evaluation_app.utils.file_utils import read_json, write_json

PROJECT_FORMAT_VERSION = 1


class ProjectError(Exception):
    """Ошибка работы с файлом проекта."""


class ProjectManager:
    """Управление сохранением, загрузкой и изменением проекта."""

    @staticmethod
    def save_project(project: Project, file_path: str) -> str:
        """Сохраняет проект в JSON-файл."""
        if not file_path:
            raise ProjectError("Не указан путь для сохранения проекта")

        data = project.to_dict()
        data["format_version"] = PROJECT_FORMAT_VERSION
        data["saved_at"] = datetime.now().isoformat(timespec="seconds")

        try:
            write_json(data, file_path)
        except OSError as exc:
            raise ProjectError(f"Не удалось сохранить проект: {exc}") from exc
        return file_path

    @staticmethod
    def load_project(file_path: str) -> Project:
        """Загружает проект из JSON-файла."""
        if not os.path.isfile(file_path):
            raise ProjectError(f"Файл проекта не найден: {file_path}")

        try:
            data = read_json(file_path)
        except ValueError as exc:
            raise ProjectError(f"Файл проекта повреждён: {exc}") from exc
        except OSError as exc:
            raise ProjectError(f"Не удалось прочитать файл проекта: {exc}") from exc

        if not isinstance(data, dict):
            raise ProjectError("Некорректный формат файла проекта")

        try:
            return Project.from_dict(data)
        except (TypeError, ValueError, AttributeError) as exc:
            raise ProjectError(f"Не удалось разобрать данные проекта: {exc}") from exc

    # --- Изменение данных преподавателя --------------------------------------

    @staticmethod
    def update_teacher(project: Project, teacher_id: str, new_teacher: Teacher) -> bool:
        """Обновляет преподавателя: старые данные удаляются, вставляются новые."""
        return project.replace_teacher(teacher_id, new_teacher)

    @staticmethod
    def replace_teacher_directions(
        project: Project, teacher_id: str, directions: Iterable
    ) -> bool:
        """Заменяет направления преподавателя."""
        return project.replace_teacher_directions(teacher_id, directions)

    @staticmethod
    def remove_teacher(project: Project, teacher_id: str) -> bool:
        """Удаляет преподавателя из проекта."""
        return project.remove_teacher(teacher_id)

    @staticmethod
    def find_teacher(project: Project, name: str) -> Teacher | None:
        """Поиск преподавателя по ФИО."""
        return project.find_teacher(name)
