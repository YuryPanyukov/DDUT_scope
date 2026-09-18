"""Работа с файлами: JSON, настройки приложения, имена файлов."""

from __future__ import annotations

import contextlib
import json
import os
import re
from datetime import datetime
from typing import Any

from evaluation_app import config

_YEAR_RE = re.compile(r"\b(19|20)\d{2}\b")
# Служебные слова в именах файлов: «Оценки_Брувель_СГ.xlsx», «Шаблон оценок.xlsx».
_SERVICE_WORDS: tuple[str, ...] = (
    "оценки", "оценка", "оценок", "оценочный", "оценочная", "лист", "ведомость",
    "протокол", "итоги", "итог", "свод", "сводка", "сводная", "таблица", "шаблон",
    "данные", "результаты", "баллы", "мониторинг", "конкурс", "конкурсы",
    "критерии", "критерий", "матрица",
)
_LEADING_SERVICE_WORDS_RE = re.compile(
    r"^(?:" + "|".join(_SERVICE_WORDS) + r")\b[\s.,:;_-]*",
    re.IGNORECASE,
)
_SERVICE_WORD_SET = {word.lower() for word in _SERVICE_WORDS}
# «Брувель СГ» → «Брувель С.Г.» (инициалы без точек в именах файлов)
_INITIALS_RE = re.compile(r"\b([А-ЯЁ])([А-ЯЁ])\b")


def ensure_dir(path: str) -> str:
    """Создаёт каталог для указанного пути (если нужно) и возвращает его."""
    directory = os.path.dirname(os.path.abspath(path))
    if directory:
        os.makedirs(directory, exist_ok=True)
    return directory


def write_json(data: Any, file_path: str, indent: int = 2) -> str:
    """Записывает данные в JSON-файл (UTF-8, без ASCII-экранирования)."""
    ensure_dir(file_path)
    with open(file_path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=indent)
    return file_path


def read_json(file_path: str) -> Any:
    """Читает JSON-файл."""
    with open(file_path, encoding="utf-8") as fh:
        return json.load(fh)


# --- Настройки приложения ----------------------------------------------------


def load_settings() -> dict:
    """Читает пользовательские настройки; при ошибке возвращает пустой словарь."""
    try:
        data = read_json(config.SETTINGS_FILE)
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def save_settings(settings: dict) -> str:
    """Сохраняет пользовательские настройки (ошибки не критичны)."""
    with contextlib.suppress(OSError):  # pragma: no cover - зависит от окружения
        write_json(settings, config.SETTINGS_FILE)
    return config.SETTINGS_FILE


def get_setting(key: str, default: Any = None) -> Any:
    """Возвращает значение настройки."""
    return load_settings().get(key, default)


def set_setting(key: str, value: Any) -> None:
    """Сохраняет одну настройку."""
    settings = load_settings()
    settings[key] = value
    save_settings(settings)


def get_default_save_path() -> str:
    """Каталог по умолчанию для проектов и экспортированных файлов."""
    path = get_setting("default_save_path") or config.DEFAULT_SAVE_PATH
    return str(path)


def set_default_save_path(path: str) -> None:
    """Запоминает каталог сохранения по умолчанию."""
    set_setting("default_save_path", str(path))


def remember_recent_file(file_path: str) -> None:
    """Добавляет файл в список недавно открытых."""
    if not file_path:
        return
    file_path = os.path.abspath(file_path)
    history = [path for path in get_recent_files() if path != file_path]
    history.insert(0, file_path)
    set_setting("recent_files", history[: config.RECENT_FILES_LIMIT])


def get_recent_files() -> list[str]:
    """Список недавно открытых проектов, которые ещё существуют на диске."""
    stored = get_setting("recent_files", [])
    if not isinstance(stored, list):
        return []
    return [str(path) for path in stored if isinstance(path, str) and os.path.isfile(path)]


# --- Имена файлов ------------------------------------------------------------


def _is_service_only(text: str) -> bool:
    """Состоит ли название только из служебных слов («шаблон оценок»)."""
    words = [word for word in re.split(r"[\s.,]+", text) if word]
    return bool(words) and all(word.lower() in _SERVICE_WORD_SET for word in words)


def timestamped_file_name(prefix: str, extension: str = "") -> str:
    """Формирует имя файла с отметкой даты и времени."""
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    name = f"{prefix}_{stamp}" if prefix else stamp
    if extension and not extension.startswith("."):
        extension = "." + extension
    return f"{name}{extension}"


def format_score(value: Any) -> str:
    """Аккуратно форматирует баллы: 85 вместо 85.0, 85.5 без изменений."""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value if value is not None else "")
    if number == int(number):
        return str(int(number))
    return f"{number:.2f}".rstrip("0").rstrip(".")


def safe_file_stem(file_path: str) -> str:
    """Превращает имя файла в название преподавателя по умолчанию.

    «Брувель_СГ.xlsx» → «Брувель С.Г.», «Оценки_Петрова_2025.xlsx» → «Петрова».
    """
    stem = os.path.splitext(os.path.basename(str(file_path)))[0]
    stem = stem.replace("_", " ").replace("-", " ")
    stem = _YEAR_RE.sub(" ", stem)
    stem = re.sub(r"\s+", " ", stem).strip(" .,")
    base = stem

    for _ in range(2):
        stripped = _LEADING_SERVICE_WORDS_RE.sub("", stem).strip(" ,")
        if stripped == stem or not stripped:
            break
        stem = stripped

    # «Шаблон оценок» → не осмысленное ФИО: возвращаем исходное имя файла.
    if _is_service_only(stem):
        stem = base

    stem = _INITIALS_RE.sub(r"\1.\2.", stem)
    # Точки в инициалах сохраняем: «Брувель С.Г.»
    return re.sub(r"\s+", " ", stem).strip(" ,")


def file_name(path: str) -> str:
    """Имя файла без каталога."""
    return os.path.basename(str(path))


def existing_path(path: str | None) -> str | None:
    """Возвращает путь, если файл существует."""
    if path and os.path.isfile(path):
        return path
    return None
