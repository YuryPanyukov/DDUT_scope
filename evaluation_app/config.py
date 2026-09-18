"""Конфигурация приложения: пути, настройки и константы."""

from __future__ import annotations

import os

APP_NAME = "Оценочный комитет"
APP_VERSION = "1.0.0"
APP_ORG = "ДДЮТ"
# Идентификатор для панели задач Windows: латиницей, без пробелов (требование Windows).
APP_USER_MODEL_ID = "ddut.scope"

# --- Пути по умолчанию -------------------------------------------------------
HOME_DIR = os.path.expanduser("~")
DEFAULT_SAVE_PATH = os.path.join(HOME_DIR, "evaluation_projects")
SETTINGS_FILE = os.path.join(HOME_DIR, ".eval_app_settings.json")

# --- Файловые фильтры --------------------------------------------------------
EXCEL_EXTENSIONS = (".xlsx", ".xlsm")
EXCEL_FILTER = "Excel files (*.xlsx *.xlsm);;All files (*)"
PROJECT_FILTER = "Проекты (*.json);;All files (*)"
PROJECT_EXTENSION = ".json"

# --- Ограничения интерфейса --------------------------------------------------
MAX_HISTORY = 20
RECENT_FILES_LIMIT = 10
PREVIEW_ROWS_LIMIT = 200
SHEET_PREVIEW_ROWS = 15
