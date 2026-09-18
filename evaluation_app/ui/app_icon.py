"""Иконка приложения: где лежит файл и как загрузить его в Qt.

Иконка нужна и окну (`setWindowIcon`), и сборке exe (PyInstaller берёт `app.ico`).
В собранном exe файлы лежат внутри каталога `_internal` (`sys._MEIPASS`), поэтому
путь считается от корня сборки, а не от расположения модуля.
"""

from __future__ import annotations

import sys
from pathlib import Path

from PyQt5.QtGui import QIcon

PACKAGE_DIR = Path(__file__).resolve().parent.parent  # каталог evaluation_app
ICONS_SUBPATH = Path("evaluation_app") / "resources" / "icons"
ICO_NAME = "app.ico"
PNG_NAME = "app.png"


def resource_base() -> Path:
    """Корень, внутри которого лежит пакет `evaluation_app`."""
    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        return Path(bundle_dir)
    return PACKAGE_DIR.parent


def icon_file(name: str = ICO_NAME) -> Path:
    """Путь к файлу иконки."""
    return resource_base() / ICONS_SUBPATH / name


def app_icon() -> QIcon:
    """Иконка приложения.

    Если файла нет (например, сборка без ресурсов), возвращается пустая иконка
    со стандартной иконкой окна — приложение из-за этого падать не должно.
    """
    for name in (ICO_NAME, PNG_NAME):
        path = icon_file(name)
        if not path.is_file():
            continue
        icon = QIcon(str(path))
        if not icon.isNull():
            return icon
    return QIcon()
