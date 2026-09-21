"""Фоновый разбор Excel-файлов с индикатором прогресса.

Загрузка и разбор книг — самая долгая операция приложения: реальные листы
преподавателей содержат тысячи строк, и синхронный разбор «подвешивал» окно.
Здесь разбор вынесен в отдельный поток (:class:`ExcelParseWorker`), а главное
окно показывает модальный прогресс-диалог до завершения.

Логика разбора остаётся в :class:`~evaluation_app.services.excel_parser.ExcelParser`;
этот модуль только оркеструет потоки, сигналы и прогресс.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass, field

from PyQt5.QtCore import QObject, Qt, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QProgressDialog, QWidget

from evaluation_app.models import Direction
from evaluation_app.services.excel_parser import ExcelParseError, ExcelParser, SheetParseResult
from evaluation_app.ui.upload_dialog import UploadSheet
from evaluation_app.utils.file_utils import file_name

ProgressCallback = Callable[[int, int, str], None]


@dataclass
class ParseOutcome:
    """Итог разбора: подготовленные листы и ошибки чтения файлов."""

    items: list[UploadSheet] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    cancelled: bool = False


def collect_upload_items(
    paths: Sequence[str],
    progress: ProgressCallback | None = None,
    is_cancelled: Callable[[], bool] | None = None,
) -> ParseOutcome:
    """Читает файлы и готовит по каждому листу элемент предпросмотра.

    Чистая функция без Qt (кроме типов ``UploadSheet``): её вызывает и фоновый
    поток, и синхронный запасной путь. ``progress(done, total, message)``
    вызывается по мере чтения файлов и разбора листов; ``total <= 0`` означает
    «неопределённый прогресс» (например, пока неизвестно число листов).
    """
    outcome = ParseOutcome()
    cancelled = is_cancelled or (lambda: False)

    def report(done: int, total: int, message: str) -> None:
        if progress is not None:
            progress(done, total, message)

    # Фаза 1: открываем книги и узнаём число листов (шаг неизвестен — busy-бар).
    loaded: list[tuple[str, ExcelParser, list[str]]] = []
    report(0, 0, "Чтение файлов…")
    for path in paths:
        if cancelled():
            outcome.cancelled = True
            return outcome
        report(0, 0, f"Чтение {file_name(path)}…")
        parser = ExcelParser(path)
        try:
            parser.load()
            sheet_names = parser.sheet_names
        except ExcelParseError as exc:
            outcome.errors.append(f"{file_name(path)}: {exc}")
            continue
        loaded.append((path, parser, list(sheet_names)))

    # Фаза 2: разбираем листы — теперь прогресс можно измерять в листах.
    total = sum(len(sheet_names) for _path, _parser, sheet_names in loaded)
    done = 0
    report(0, total, "Разбор листов…")
    for path, parser, sheet_names in loaded:
        for sheet_name in sheet_names:
            if cancelled():
                outcome.cancelled = True
                return outcome
            report(done, total, f"Разбор {file_name(path)} / {sheet_name}…")
            item = UploadSheet(
                file_path=path,
                sheet_name=sheet_name,
                parser=parser,
                default_teacher=parser.default_teacher_name(),
            )
            try:
                item.result = parser.parse_sheet(sheet_name)
            except ExcelParseError as exc:
                item.error = str(exc)
                item.result = SheetParseResult(
                    sheet_name=sheet_name,
                    direction=Direction(name=sheet_name),
                )
            outcome.items.append(item)
            done += 1
            report(done, total, f"Разбор {file_name(path)} / {sheet_name}…")

    report(total, total, "Готово")
    return outcome


class ExcelParseWorker(QObject):
    """Выполняет :func:`collect_upload_items` в отдельном потоке."""

    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal(object)  # ParseOutcome

    def __init__(self, paths: Sequence[str]) -> None:
        super().__init__()
        self._paths = list(paths)
        self._cancelled = False

    def cancel(self) -> None:
        """Просит прервать разбор: флаг проверяется между файлами и листами."""
        self._cancelled = True

    @pyqtSlot()
    def run(self) -> None:
        """Точка входа потока."""
        try:
            outcome = collect_upload_items(
                self._paths,
                progress=self.progress.emit,
                is_cancelled=lambda: self._cancelled,
            )
        except Exception as exc:  # поток не должен «умирать» молча
            outcome = ParseOutcome(errors=[f"Непредвиденная ошибка разбора: {exc}"])
        self.finished.emit(outcome)


class _ParseDialogController(QObject):
    """Связывает сигналы воркера с прогресс-диалогом (в главном потоке)."""

    def __init__(self, dialog: QProgressDialog) -> None:
        super().__init__(dialog)
        self._dialog = dialog
        self.outcome: ParseOutcome | None = None

    @pyqtSlot(int, int, str)
    def on_progress(self, done: int, total: int, message: str) -> None:
        """Обновляет прогресс-бар и подпись."""
        if total > 0:
            self._dialog.setRange(0, total)
            self._dialog.setValue(min(done, total))
        else:
            self._dialog.setRange(0, 0)
        self._dialog.setLabelText(message)

    @pyqtSlot(object)
    def on_finished(self, outcome: ParseOutcome) -> None:
        """Запоминает итог и закрывает диалог."""
        self.outcome = outcome
        self._dialog.setValue(self._dialog.maximum())
        self._dialog.accept()


def run_parse(paths: Sequence[str], parent: QWidget | None = None) -> ParseOutcome | None:
    """Разбирает файлы в фоне, показывая модальный прогресс-диалог.

    Возвращает итог разбора или ``None``, если пользователь отменил операцию.
    """
    worker = ExcelParseWorker(paths)
    thread = QThread()
    worker.moveToThread(thread)

    dialog = QProgressDialog("Подготовка…", "Отмена", 0, 0, parent)
    dialog.setWindowTitle("Чтение Excel-файлов")
    dialog.setWindowModality(Qt.WindowModal)
    dialog.setMinimumDuration(0)
    dialog.setAutoClose(False)
    dialog.setAutoReset(False)
    dialog.setWindowFlags(
        (dialog.windowFlags() | Qt.CustomizeWindowHint) & ~Qt.WindowCloseButtonHint
    )
    dialog.setMinimumWidth(420)

    controller = _ParseDialogController(dialog)

    worker.progress.connect(controller.on_progress)
    worker.finished.connect(controller.on_finished)
    worker.finished.connect(thread.quit)
    # Прямое соединение: отмена должна выставить флаг сразу, а не ждать
    # окончания цикла событий занятого воркер-потока.
    dialog.canceled.connect(worker.cancel, type=Qt.DirectConnection)

    thread.started.connect(worker.run)
    thread.start()
    dialog.exec_()
    # Если диалог закрыли кнопкой «Отмена», воркер ещё дорабатывает — дожидаемся
    # остановки потока, иначе QThread будет уничтожен во время выполнения.
    if thread.isRunning():
        thread.quit()
    thread.wait()

    outcome = controller.outcome
    if outcome is not None and outcome.cancelled:
        # Разбор прерван: освобождаем открытые книги, чтобы файлы Excel
        # не остались заблокированными (актуально для Windows).
        for item in outcome.items:
            if item.parser is not None:
                item.parser.close()
        return None
    return outcome
