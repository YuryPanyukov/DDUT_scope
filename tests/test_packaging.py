"""Согласованность упаковки: версия, точка входа и группа сборки.

Версия приложения показывается в окне (`config.APP_VERSION`) и попадает в имя
архива при сборке exe — расхождение с `pyproject.toml` заметно пользователю
как «разные версии» у одного и того же файла.
"""

from __future__ import annotations

import struct
import tomllib
from pathlib import Path

from evaluation_app import config

ROOT = Path(__file__).resolve().parent.parent


def project_config() -> dict:
    """Содержимое pyproject.toml."""
    return tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))


def test_app_version_matches_pyproject() -> None:
    assert project_config()["project"]["version"] == config.APP_VERSION


def test_console_script_points_to_main() -> None:
    scripts = project_config()["project"]["scripts"]
    assert scripts["main"] == "evaluation_app.main:main"


def test_build_group_has_pyinstaller() -> None:
    """CI собирает exe этой группой — её нельзя потерять при правке зависимостей."""
    groups = project_config()["dependency-groups"]
    assert any(dependency.startswith("pyinstaller") for dependency in groups["build"])


def test_build_script_exists() -> None:
    assert (ROOT / "scripts" / "build_exe.py").is_file()


def test_app_icon_files_are_committed() -> None:
    """Иконка нужна и сборке exe, и работающему приложению — файлы лежат в репозитории."""
    icons = ROOT / "evaluation_app" / "resources" / "icons"
    ico, png = icons / "app.ico", icons / "app.png"

    assert ico.stat().st_size > 0
    assert png.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"

    # ICO: зарезервировано, тип 1 (иконка), и несколько размеров внутри.
    reserved, kind, count = struct.unpack("<HHH", ico.read_bytes()[:6])
    assert (reserved, kind) == (0, 1)
    assert count >= 4
