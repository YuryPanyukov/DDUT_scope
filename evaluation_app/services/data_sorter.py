"""Сортировка агрегированных данных (см. п. 4.2 плана)."""

from __future__ import annotations

from collections.abc import Iterable

from evaluation_app.utils.validators import name_key

# Ключ сортировки → подпись для интерфейса.
SORT_KEYS: dict[str, str] = {
    "name": "ФИО (А-Я)",
    "score": "Итого баллов",
    "competitions": "Число конкурсов",
    "entries": "Число оценок",
}


class DataSorter:
    """Сортировка сводных записей."""

    @staticmethod
    def sort_by_name(data: Iterable[dict], reverse: bool = False) -> list[dict]:
        """Сортировка по ФИО (по алфавиту, без учёта регистра)."""
        return sorted(data, key=lambda item: name_key(item.get("student_name", "")), reverse=reverse)

    @staticmethod
    def sort_by_score(data: Iterable[dict], reverse: bool = True) -> list[dict]:
        """Сортировка по баллам (по убыванию по умолчанию)."""
        return sorted(
            data,
            key=lambda item: (
                float(item.get("total_score") or 0.0),
                name_key(item.get("student_name", "")),
            ),
            reverse=reverse,
        )

    @staticmethod
    def sort_by_competitions(data: Iterable[dict], reverse: bool = True) -> list[dict]:
        """Сортировка по количеству конкурсов."""
        return sorted(
            data,
            key=lambda item: (
                len(item.get("competitions") or []),
                name_key(item.get("student_name", "")),
            ),
            reverse=reverse,
        )

    @staticmethod
    def sort_by_entries(data: Iterable[dict], reverse: bool = True) -> list[dict]:
        """Сортировка по количеству оценок."""
        return sorted(
            data,
            key=lambda item: (
                int(item.get("entries") or 0),
                name_key(item.get("student_name", "")),
            ),
            reverse=reverse,
        )

    @staticmethod
    def sort(data: Iterable[dict], key: str = "score", reverse: bool = True) -> list[dict]:
        """Универсальная сортировка по ключу из :data:`SORT_KEYS`."""
        handlers = {
            "name": DataSorter.sort_by_name,
            "score": DataSorter.sort_by_score,
            "competitions": DataSorter.sort_by_competitions,
            "entries": DataSorter.sort_by_entries,
        }
        handler = handlers.get(key, DataSorter.sort_by_score)
        return handler(data, reverse=reverse)

    @staticmethod
    def default_reverse(key: str) -> bool:
        """Направление сортировки по умолчанию для ключа."""
        return key != "name"
