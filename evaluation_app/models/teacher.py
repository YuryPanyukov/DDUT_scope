"""Модель преподавателя."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, field
from uuid import uuid4

from evaluation_app.models.direction import Direction
from evaluation_app.utils.validators import name_key, normalize_name


@dataclass
class Teacher:
    """Преподаватель и его направления (листы Excel)."""

    id: str = field(default_factory=lambda: str(uuid4()))
    name: str = ""
    directions: list[Direction] = field(default_factory=list)

    def __post_init__(self) -> None:
        self.name = normalize_name(self.name)

    # --- Направления ---------------------------------------------------------

    def add_direction(
        self,
        direction: Direction | None = None,
        *,
        name: str = "",
        replace: bool = True,
    ) -> Direction:
        """Добавляет направление; при совпадении имени заменяет данные."""
        if direction is None:
            direction = Direction(name=name)

        if replace:
            existing = self.get_direction(direction.name)
            if existing is not None:
                existing.replace_evaluations(direction.evaluations)
                return existing

        self.directions.append(direction)
        return direction

    def remove_direction(self, direction_id: str) -> bool:
        """Удаляет направление по идентификатору."""
        for index, direction in enumerate(self.directions):
            if direction.id == direction_id:
                del self.directions[index]
                return True
        return False

    def get_direction(self, name: str) -> Direction | None:
        """Находит направление по имени (без учёта регистра)."""
        key = name_key(name)
        for direction in self.directions:
            if name_key(direction.name) == key:
                return direction
        return None

    def get_direction_by_id(self, direction_id: str) -> Direction | None:
        """Находит направление по идентификатору."""
        for direction in self.directions:
            if direction.id == direction_id:
                return direction
        return None

    def replace_directions(self, directions: Iterable[Direction]) -> None:
        """Полностью заменяет направления (используется при обновлении из Excel)."""
        self.directions = list(directions)

    # --- Агрегаты ------------------------------------------------------------

    def evaluations(self) -> list:
        """Все оценки преподавателя по всем направлениям."""
        return [
            evaluation
            for direction in self.directions
            for evaluation in direction.evaluations
        ]

    def total_score(self) -> float:
        """Сумма баллов преподавателя."""
        return float(sum(direction.total_score() for direction in self.directions))

    def student_names(self) -> list[str]:
        """Уникальные ФИО учеников преподавателя."""
        result: list[str] = []
        for evaluation in self.evaluations():
            if evaluation.student_name and evaluation.student_name not in result:
                result.append(evaluation.student_name)
        return result

    def matches_name(self, name: str) -> bool:
        """Сравнение ФИО без учёта регистра и лишних пробелов."""
        return name_key(self.name) == name_key(name)

    def to_dict(self) -> dict:
        """Сериализация в словарь."""
        return {
            "id": self.id,
            "name": self.name,
            "directions": [direction.to_dict() for direction in self.directions],
        }

    @classmethod
    def from_dict(cls, data: dict) -> Teacher:
        """Восстановление из словаря."""
        data = data or {}
        teacher = cls(
            id=str(data.get("id") or uuid4()),
            name=str(data.get("name") or ""),
        )
        for item in data.get("directions") or []:
            if isinstance(item, dict):
                teacher.directions.append(Direction.from_dict(item))
        return teacher
