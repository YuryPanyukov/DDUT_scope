"""Сборка exe-версии приложения для Windows через PyInstaller.

Запуск:

    uv run --group build python scripts/build_exe.py

Результат в каталоге ``dist``:

* ``dist/DDUT-Scope/DDUT-Scope.exe`` — приложение одной папкой (запускается
  быстрее, чем onefile, и реже попадает под ложные срабатывания антивирусов);
* ``dist/DDUT-Scope-<версия>-windows-x64.zip`` — та же папка с краткой
  инструкцией, готовая к передаче.

Релизный workflow ``.github/workflows/release.yml`` вызывает скрипт с опцией
``--print-version``, чтобы сверить тег с версией и назвать релиз: опция просто
печатает версию из pyproject.toml, ничего не собирая.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tomllib
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

APP_NAME = "DDUT-Scope"
ENTRY_POINT = ROOT / "main.py"
ICONS_DIR = ROOT / "evaluation_app" / "resources" / "icons"
ICON = ICONS_DIR / "app.ico"
# Куда положить ресурсы внутри сборки: путь совпадает с sys._MEIPASS-путём в app_icon.py.
ICONS_DEST = "evaluation_app/resources/icons"
DIST_DIR = ROOT / "dist"
BUILD_DIR = ROOT / "build"
APP_DIR = DIST_DIR / APP_NAME

INSTRUCTION_FILE = "Как запускать.txt"
INSTRUCTION = """Оценочный комитет (DDUT-Scope)
===============================

1. Распакуйте архив целиком в любую папку, например на рабочий стол.
   Запускать DDUT-Scope.exe прямо из архива нельзя — приложение не найдёт
   свои файлы.

2. Запустите DDUT-Scope.exe.

3. Данные приложения сохраняются в файл проекта (JSON) — его можно открыть
   позже через меню приложения.

Личные файлы преподавателей (Excel) на обработку не влияют: они не входят
в поставку и не нужны для запуска.
"""


def project_version() -> str:
    """Версия проекта из pyproject.toml — она же попадает в имя архива."""
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    return str(data["project"]["version"])


def pyinstaller_command() -> list[str]:
    """Команда сборки одной папкой без консольного окна."""
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        "--windowed",
        "--name",
        APP_NAME,
        "--distpath",
        str(DIST_DIR),
        "--workpath",
        str(BUILD_DIR / "work"),
        "--specpath",
        str(BUILD_DIR / "spec"),
        # Корень проекта в путях поиска: пакет evaluation_app может быть не установлен.
        "--paths",
        str(ROOT),
    ]
    if ICON.is_file():
        command += ["--icon", str(ICON)]
    if ICONS_DIR.is_dir():
        # Иконка нужна приложению и во время работы: setWindowIcon читает файл с диска.
        command += ["--add-data", f"{ICONS_DIR}{os.pathsep}{ICONS_DEST}"]
    command.append(str(ENTRY_POINT))
    return command


def build() -> Path:
    """Собирает приложение и возвращает путь к exe."""
    if APP_DIR.exists():
        shutil.rmtree(APP_DIR)

    print("Сборка:", " ".join(pyinstaller_command()))
    subprocess.run(pyinstaller_command(), check=True, cwd=ROOT)

    exe_path = APP_DIR / f"{APP_NAME}.exe"
    if not exe_path.is_file():
        raise SystemExit(f"Сборка завершилась без файла {exe_path}")
    return exe_path


def make_archive(version: str) -> Path:
    """Упаковывает собранную папку в zip вместе с инструкцией."""
    (APP_DIR / INSTRUCTION_FILE).write_text(INSTRUCTION, encoding="utf-8")

    archive = DIST_DIR / f"{APP_NAME}-{version}-windows-x64.zip"
    archive.unlink(missing_ok=True)
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(APP_DIR.rglob("*")):
            if path.is_file():
                zf.write(path, Path(APP_NAME) / path.relative_to(APP_DIR))
    return archive


def folder_size(path: Path) -> float:
    """Размер каталога в мегабайтах."""
    total = sum(item.stat().st_size for item in path.rglob("*") if item.is_file())
    return total / (1024 * 1024)


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if "--print-version" in args:
        print(project_version())
        return 0

    if sys.platform != "win32":
        print("Сборка exe рассчитана на Windows: PyInstaller не кросс-компилятор.")
        return 1

    exe_path = build()
    version = project_version()
    archive = make_archive(version)

    print()
    print(f"Приложение: {exe_path} ({folder_size(APP_DIR):.1f} МБ)")
    print(f"Архив:      {archive} ({archive.stat().st_size / (1024 * 1024):.1f} МБ)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
