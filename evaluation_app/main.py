"""Точка входа приложения «Оценочный комитет»."""

from __future__ import annotations

import sys
from collections.abc import Sequence

from PyQt5.QtWidgets import QApplication

from evaluation_app import config
from evaluation_app.ui.app_icon import app_icon


def _set_windows_app_id() -> None:
    """Идентификатор приложения для панели задач Windows.

    Без него Windows группирует окно со значком интерпретатора, а не с иконкой
    приложения. Значение должно быть без пробелов и коротким — см. APP_USER_MODEL_ID.
    """
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(config.APP_USER_MODEL_ID)
    except Exception:  # влияет только на вид в панели задач
        return


def main(argv: Sequence[str] | None = None) -> int:
    """Запускает GUI-приложение."""
    arguments = list(argv) if argv is not None else list(sys.argv)

    app = QApplication(arguments)
    app.setApplicationName(config.APP_NAME)
    app.setApplicationVersion(config.APP_VERSION)
    app.setOrganizationName(config.APP_ORG)

    _set_windows_app_id()
    icon = app_icon()
    if not icon.isNull():
        # Иконка приложения: её получают и диалоги, и окно, и панель задач.
        app.setWindowIcon(icon)

    # Импорт внутри функции: позволяет запускать CLI-утилиты без GUI.
    from evaluation_app.ui.main_window import MainWindow

    window = MainWindow()
    window.show()
    return app.exec_()


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
