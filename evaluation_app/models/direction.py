"""Модель направления (обычно — лист Excel)."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, field
from uuid import uuid4

from evaluation_app.models.evaluation import Evaluation


@dataclass
class Direction:
    """Направление преподавателя со списком оценок."""

    id: str = field(default_factory=lambda: str(uuid4()))
    name: str = ""
    evaluations: list[Evaluation] = field(default_factory=list)

    def add_evaluation(
        self,
        evaluation: Evaluation | None = None,
        *,
        student_name: str = "",
        competition_name: str = "",
        total_score: float = 0.0,
    ) -> Evaluation:
        """Добавляет оценку и возвращает её."""
        if evaluation is None:
            evaluation = Evaluation(
                student_name=student_name,
                competition_name=competition_name,
                total_score=total_score,
            )
        self.evaluations.append(evaluation)
        return evaluation

    def remove_evaluation(self, evaluation_id: str) -> bool:
        """Удаляет оценку по идентификатору; ``False``, если её не было.

        Используется только в тестах: при обновлении из Excel оценки
        заменяются вместе с направлением целиком (:meth:`replace_evaluations`).
        """
        for index, evaluation in enumerate(self.evaluations):
            if evaluation.id == evaluation_id:
                del self.evaluations[index]
                return True
        return False

    def get_evaluation(self, evaluation_id: str) -> Evaluation | None:
        """Возвращает оценку по идентификатору или ``None``.

        Используется только в тестах: интерфейс работает с готовыми списками
        оценок и поштучно их не ищет.
        """
        for evaluation in self.evaluations:
            if evaluation.id == evaluation_id:
                return evaluation
        return None

    def replace_evaluations(self, evaluations: Iterable[Evaluation]) -> None:
        """Полностью заменяет список оценок."""
        self.evaluations = list(evaluations)

    def total_score(self) -> float:
        """Сумма баллов по направлению."""
        return float(sum(evaluation.total_score for evaluation in self.evaluations))

    def student_names(self) -> list[str]:
        """Уникальные ФИО учеников направления в порядке появления.

        Используется только в тестах: подписи в интерфейсе считаются на уровне
        преподавателя (:meth:`Teacher.student_names`) и в агрегатах
        (:meth:`DataMerger.totals_by_teacher`).
        """
        result: list[str] = []
        for evaluation in self.evaluations:
            if evaluation.student_name and evaluation.student_name not in result:
                result.append(evaluation.student_name)
        return result

    def __len__(self) -> int:
        """Число оценок в направлении.

        Используется только в тестах (``len(direction)``): в коде длина берётся
        у ``direction.evaluations``.
        """
        return len(self.evaluations)

    def to_dict(self) -> dict:
        """Сериализация в словарь."""
        return {
            "id": self.id,
            "name": self.name,
            "evaluations": [evaluation.to_dict() for evaluation in self.evaluations],
        }

    @classmethod
    def from_dict(cls, data: dict) -> Direction:
        """Восстановление из словаря."""
        data = data or {}
        direction = cls(
            id=str(data.get("id") or uuid4()),
            name=str(data.get("name") or ""),
        )
        for item in data.get("evaluations") or []:
            if isinstance(item, dict):
                direction.evaluations.append(Evaluation.from_dict(item))
        return direction
