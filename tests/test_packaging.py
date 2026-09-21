"""Согласованность упаковки: версия, точка входа, группа сборки и порог vulture.

Версия приложения показывается в окне (`config.APP_VERSION`) и попадает в имя
архива при сборке exe — расхождение с `pyproject.toml` заметно пользователю
как «разные версии» у одного и того же файла. Настройки vulture тоже здесь:
ослабленный порог, шаблон в его исключениях или забытая пометка «Используется
только в тестах» тихо пропускают мёртвый код через CI.
"""

from __future__ import annotations

import ast
import struct
import tomllib
from pathlib import Path

import pytest

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


def vulture_config() -> dict:
    """Секция [tool.vulture] из pyproject.toml."""
    return project_config()["tool"]["vulture"]


def test_dead_code_threshold_is_minimal() -> None:
    """Порог vulture — минимальный, иначе новые мёртвые функции пройдут CI.

    vulture сообщает о неиспользуемых функциях, классах, методах и свойствах
    ровно с уверенностью 60%: подняв порог, проверка перестанет их замечать.
    """
    assert vulture_config()["min_confidence"] <= 60


def test_dead_code_scan_covers_package() -> None:
    """Сканируется сам пакет, а не только тесты или скрипты."""
    assert "evaluation_app" in vulture_config()["paths"]


def test_dead_code_ignores_are_precise() -> None:
    """В ignore_names нет шаблонов: иначе они скрыли бы и новые определения."""
    for name in vulture_config()["ignore_names"]:
        assert not any(char in name for char in "*?["), f"{name!r} скроет новые мёртвые определения"


TEST_ONLY_MARKER = "Используется только в тестах"


def _unused_package_definitions() -> tuple[list, int]:
    """Определения пакета, на которые нет ссылок вне тестов.

    :return: список таких определений (функции, методы и свойства) и число
        разобранных определений — второе показывает, что скан состоялся.
    """
    vulture = pytest.importorskip("vulture")
    with (ROOT / "pyproject.toml").open("rb") as toml_file:
        settings = vulture.config.make_config(argv=[], tomlfile=toml_file)

    checker = vulture.Vulture(ignore_names=settings["ignore_names"])
    checker.scavenge([str(ROOT / "evaluation_app")], exclude=settings["exclude"])
    unused = checker.get_unused_code(min_confidence=settings["min_confidence"])
    callables = [item for item in unused if item.typ in ("function", "method", "property")]
    parsed = len(checker.defined_funcs) + len(checker.defined_methods)
    return callables, parsed


def _definitions_marked_as_test_only() -> set[tuple[Path, str]]:
    """Определения, в docstring которых есть пометка «используется только в тестах»."""
    marked: set[tuple[Path, str]] = set()
    for path in (ROOT / "evaluation_app").rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                continue
            if TEST_ONLY_MARKER in (ast.get_docstring(node) or ""):
                marked.add((path.resolve(), node.name))
    return marked


def test_definitions_without_callers_are_marked_as_test_only() -> None:
    """У определения без вызовов в приложении должна быть пометка в docstring.

    Основной прогон vulture идёт вместе с тестами и поэтому молчит о методах,
    которые вызывает только тестовый код. Здесь тот же скан повторяется без
    тестов: если метод не вызывается из приложения и при этом не помечен как
    тестовый, значит он либо забыт (настоящий мёртвый код), либо это новый
    тестовый API без пояснения, зачем он нужен. vulture сопоставляет имена,
    поэтому метод, который вызывается только из другого непомеченного метода,
    он не заметит — проверка заведомо неполная, но регрессии не пропустит.
    """
    unused, parsed = _unused_package_definitions()
    assert parsed > 0, "vulture не разобрал пакет — проверка выродилась"

    marked = _definitions_marked_as_test_only()
    unmarked = [
        f"{Path(item.filename).relative_to(ROOT)}:{item.first_lineno} {item.name}"
        for item in unused
        if (Path(item.filename).resolve(), item.name) not in marked
    ]
    assert not unmarked, f"нет вызовов в приложении и нет пометки «{TEST_ONLY_MARKER}»: " + ", ".join(
        unmarked
    )


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
