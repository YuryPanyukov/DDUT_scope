"""Модель одной оценки ученика."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any
from uuid import uuid4

from evaluation_app.utils.validators import normalize_name, parse_score


@dataclass
class Evaluation:
    """Оценка ученика по одному конкурсу внутри направления."""

    id: str = field(default_factory=lambda: str(uuid4()))
    student_name: str = ""
    competition_name: str = ""
    total_score: float = 0.0

    def __post_init__(self) -> None:
        self.student_name = normalize_name(self.student_name)
        self.competition_name = normalize_name(self.competition_name)
        score = parse_score(self.total_score)
        self.total_score = float(score) if score is not None else 0.0

    def to_dict(self) -> dict:
        """Сериализация в словарь (для JSON-проекта)."""
        return {
            "id": self.id,
            "student_name": self.student_name,
            "competition_name": self.competition_name,
            "total_score": self.total_score,
        }

    @classmethod
    def from_dict(cls, data: dict) -> Evaluation:
        """Восстановление из словаря."""
        data = data or {}
        return cls(
            id=str(data.get("id") or uuid4()),
            student_name=str(data.get("student_name") or ""),
            competition_name=str(data.get("competition_name") or ""),
            total_score=data.get("total_score") or 0.0,
        )

    def __str__(self) -> str:  # pragma: no cover - для отладки
        """Читаемое представление «ФИО — конкурс: балл».

        Нужно только для отладки и сообщений в логе; в интерфейсе не
        используется (там поля выводятся по отдельности).
        """
        return f"{self.student_name} — {self.competition_name}: {self.total_score}"

    def __eq__(self, other: Any) -> bool:
        """Равенство по содержимому: сравнивает сериализованные поля.

        Идентификатор не учитывается, поэтому восстановленная из файла оценка
        равна исходной; так сравниваются оценки как значения, а не как записи
        с одинаковым ``id``. Определяя ``__eq__`` вручную, объект теряет
        сгенерированный ``__hash__`` (становится ``None``) — оценки нельзя
        добавлять в ``set`` или использовать как ключи ``dict``.
        """
        if not isinstance(other, Evaluation):
            return NotImplemented
        return self.to_dict() == other.to_dict()
