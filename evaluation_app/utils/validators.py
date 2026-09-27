"""Валидация данных и разбор значений, полученных из Excel.

Здесь собрана «гибкая» логика определения столбцов: заголовки в файлах
преподавателей могут называться по-разному («ФИО» / «Ученик» /
«Фамилия Имя Отчество»), поэтому используется сопоставление по синонимам.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping
from typing import Any

# Возможные названия столбцов в файлах преподавателей.
COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "student_name": (
        "фио",
        "ф.и.о",
        "фио ученика",
        "фио участника",
        "фамилия имя отчество",
        "фамилия и.о",
        "фамилия и имя",
        "ученик",
        "ученики",
        "обучающийся",
        "обучающаяся",
        "участник",
        "воспитанник",
    ),
    "competition_name": (
        "название конкурса",
        "название мероприятия",
        "конкурс",
        "мероприятие",
        "номинация",
        "название",
        "вид конкурса",
    ),
    "total_score": (
        "итого баллов",
        "итого",
        "итог",
        "итоги",
        "сумма баллов",
        "сумма",
        "баллы",
        "балл",
        "общий балл",
        "результат",
        "всего",
        "total",
        "score",
    ),
}

# Служебные столбцы листа-матрицы: они не являются критериями (например,
# «за вход в систему» в реальных оценочных листах ДДЮТ). Баллы столбца входа
# парсер прибавляет к итогу ученика отдельно — см. :func:`is_login_column`.
# Шаблоны подобраны с учётом границ слова: «мастер-классах» — критерий,
# а столбец «класс» — служебный.
NON_CRITERION_LABEL_PATTERNS: tuple[str, ...] = (
    r"вход\s+в\s+систему",
    r"^№",
    r"(?<![\w-])логин\w*",
    r"(?<![\w-])номер(?!\w)",
    r"(?<![\w-])(класс|классы)(?![\w-])",
    r"(?<![\w-])(группа|группы|группе|группу)(?![\w-])",
    r"(?<![\w-])педагог\w*",
    r"(?<![\w-])преподавател\w*",
    r"(?<![\w-])руководител\w*",
    r"(?<![\w-])комментар\w*",
    r"(?<![\w-])примечани\w*",
    r"(?<![\w-])дата(?!\w)",
    r"(?<![\w-])статус\w*",
    r"(?<![\w-])фото\w*",
    r"(?<![\w-])user\w*",
    r"(?<![\w-])фио(?![\w-])",
)
_NON_CRITERION_RE = re.compile("|".join(NON_CRITERION_LABEL_PATTERNS))
# Служебный столбец с баллами входа: формула ИТОГО в реальных файлах
# (=SUM(C:AF)) его не учитывает, поэтому парсер прибавляет его баллы сам.
LOGIN_COLUMN_RE = re.compile(r"вход\s+в\s+систему")
# Подписи из одних значений («30», «40/50/60», «90%») — это максимумы, а не названия.
_VALUE_ONLY_RE = re.compile(r"^[\d\s.,%/№-]+$")

# Строки-маркеры итогов: не должны попадать в список учеников.
TOTAL_MARKERS: tuple[str, ...] = ("итого", "итог", "итоги", "всего", "сумма", "total")

FIELD_TITLES: dict[str, str] = {
    "student_name": "ФИО ученика",
    "competition_name": "Название конкурса",
    "total_score": "Итого баллов",
}

# Поля, без которых строку нельзя разобрать осмысленно.
REQUIRED_FIELDS: tuple[str, ...] = ("student_name", "total_score")

_NUMBER_RE = re.compile(r"-?\d+(?:[.,]\d+)?")
_SPACES_RE = re.compile(r"\s+")


def normalize_header(value: Any) -> str:
    """Приводит заголовок столбца к сравнимому виду."""
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").replace("\n", " ")
    text = text.replace("ё", "е").replace("Ё", "Е")
    return _SPACES_RE.sub(" ", text.strip().lower())


def normalize_name(value: Any) -> str:
    """Убирает лишние пробелы в ФИО, сохраняя регистр."""
    if value is None:
        return ""
    return _SPACES_RE.sub(" ", str(value).replace("\xa0", " ")).strip()


def name_key(value: Any) -> str:
    """Ключ для объединения записей об одном ученике."""
    return normalize_name(value).lower()


def match_column(header: Any) -> str | None:
    """Определяет, какому полю соответствует заголовок столбца."""
    norm = normalize_header(header)
    if not norm:
        return None

    # Сначала ищем точное совпадение с синонимом.
    for field, aliases in COLUMN_ALIASES.items():
        if norm in aliases:
            return field

    # Затем — вхождение синонима в заголовок.
    for field, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in norm:
                return field
    return None


def detect_columns(headers: Iterable[Any]) -> dict[str, int]:
    """Сопоставляет поля и номера столбцов (нумерация с 1)."""
    mapping: dict[str, int] = {}
    for index, header in enumerate(headers, start=1):
        field = match_column(header)
        if field and field not in mapping:
            mapping[field] = index
    return mapping


def exact_match_column(header: Any) -> str | None:
    """Строгое сопоставление заголовка с синонимом (без вхождений).

    Нужно для листов-матриц: там критерии называются «Участие в мероприятиях
    района», «Разовое мероприятие» и т. п. — нечёткое сравнение ошибочно
    приняло бы их за столбец «Название конкурса».
    """
    norm = normalize_header(header)
    if not norm:
        return None
    for field, aliases in COLUMN_ALIASES.items():
        if norm in aliases:
            return field
    return None


def is_service_label(label: Any) -> bool:
    """Является ли подпись столбца служебной (не критерием)."""
    norm = normalize_header(label)
    if not norm:
        return True
    if _VALUE_ONLY_RE.match(norm):
        return True  # нумерация столбцов или строка максимумов
    return _NON_CRITERION_RE.search(norm) is not None


def is_login_column(label: Any) -> bool:
    """Служебный ли это столбец с баллами входа («за вход в систему»)."""
    return LOGIN_COLUMN_RE.search(normalize_header(label)) is not None


def is_total_marker(value: Any) -> bool:
    """Похожа ли строка на итоговую (а не на ФИО ученика)."""
    norm = normalize_header(value)
    if not norm:
        return False
    return any(norm.startswith(marker) for marker in TOTAL_MARKERS)


def is_header_or_total_marker(value: Any) -> bool:
    """Строка похожа на повторный заголовок или итоговую строку.

    Такие строки встречаются в реальных файлах (заголовок повторяется на
    каждой странице, снизу пишут «ИТОГО»), и учениками они не являются.
    """
    norm = normalize_header(value)
    if not norm:
        return False
    if is_total_marker(norm):
        return True
    return norm in COLUMN_ALIASES["student_name"]


def missing_fields(mapping: Mapping[str, Any]) -> list[str]:
    """Возвращает список отсутствующих обязательных полей."""
    return [field for field in REQUIRED_FIELDS if field not in mapping]


def parse_score(value: Any) -> float | None:
    """Разбирает баллы из ячейки Excel. Возвращает None, если не число."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).replace("\xa0", "").replace(" ", "").strip()
    if not text or text.startswith("="):
        # Формула без вычисленного значения не является баллом.
        return None
    text = text.replace(",", ".")

    match = _NUMBER_RE.search(text)
    if not match:
        return None
    try:
        return float(match.group().replace(",", "."))
    except ValueError:  # pragma: no cover - защита от неожиданных значений
        return None


def detect_columns_exact(headers: Iterable[Any]) -> dict[str, int]:
    """Сопоставление полей только по точным синонимам заголовков.

    Нужно там, где ошибочное «угадывание» дороже пропуска: в листах-матрицах
    группа «УЧЕБНАЯ... (КОНКУРСЫ...)» не должна становиться столбцом «Название
    конкурса», иначе столбец выпадет из критериев.
    """
    mapping: dict[str, int] = {}
    for index, header in enumerate(headers, start=1):
        field = exact_match_column(header)
        if field and field not in mapping:
            mapping[field] = index
    return mapping


def shorten(value: Any, limit: int = 24) -> str:
    """Обрезает длинные подписи для интерфейса."""
    text = normalize_name(value)
    if len(text) <= limit:
        return text
    return text[:limit].rstrip(" ,.;") + "…"


def headers_from(source: Any, row: int = 1) -> list[Any]:
    """Достаёт заголовки из листа Excel, списка или словаря.

    Используется только в тестах — через
    :meth:`DataValidator.validate_excel_headers`.
    """
    if source is None:
        return []
    if hasattr(source, "cell") and hasattr(source, "max_column"):
        column_count = source.max_column or 0
        return [source.cell(row=row, column=col).value for col in range(1, column_count + 1)]
    if isinstance(source, Mapping):
        return list(source.values())
    return list(source)


class DataValidator:
    """Валидация заголовков Excel-таблиц и значений баллов (см. этап 7 плана)."""

    @staticmethod
    def validate_excel_headers(source: Any, row: int = 1) -> bool:
        """Проверяет, что в таблице есть обязательные столбцы.

        Принимает «таблицу» в любом виде (лист openpyxl, список, словарь)
        и не зависит от модели проекта. Используется только в тестах: при
        разборе столбцы определяет сам парсер (:func:`detect_columns`).
        """
        mapping = detect_columns(headers_from(source, row=row))
        return not missing_fields(mapping)

    @staticmethod
    def validate_score(value: Any) -> bool:
        """Проверяет, что значение баллов — число.

        Используется только в тестах: при разборе баллы приводит
        :func:`parse_score`, который сам отсеивает нечисловые значения.
        """
        return parse_score(value) is not None

    @staticmethod
    def describe_missing(mapping: Mapping[str, Any]) -> str:
        """Человекочитаемое описание недостающих столбцов."""
        missing = missing_fields(mapping)
        if not missing:
            return ""
        titles = ", ".join(FIELD_TITLES.get(field, field) for field in missing)
        return f"Не найдены обязательные столбцы: {titles}"
