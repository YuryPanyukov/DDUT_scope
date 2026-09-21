"""Главное окно приложения «Оценочный комитет» (см. п. 5.1 плана)."""

from __future__ import annotations

import os
from collections.abc import Sequence

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QAction,
    QDialog,
    QFileDialog,
    QInputDialog,
    QLabel,
    QMainWindow,
    QMessageBox,
    QSplitter,
    QStyle,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from evaluation_app import config
from evaluation_app.models import Direction, Project
from evaluation_app.services.data_merger import DataMerger
from evaluation_app.services.data_sorter import DataSorter
from evaluation_app.services.excel_export import ExcelExporter
from evaluation_app.services.excel_parser import ExcelParseError, ExcelParser
from evaluation_app.services.project_manager import ProjectError, ProjectManager
from evaluation_app.ui.app_icon import app_icon
from evaluation_app.ui.direction_view import DirectionView
from evaluation_app.ui.parse_worker import collect_upload_items, run_parse
from evaluation_app.ui.settings_dialog import AppSettingsDialog, ProjectSettingsDialog, default_academic_year
from evaluation_app.ui.summary_table_view import SummaryTableView
from evaluation_app.ui.teacher_list_widget import TeacherListWidget
from evaluation_app.ui.upload_dialog import (
    ColumnMappingDialog,
    UploadPreviewDialog,
    UploadSheet,
)
from evaluation_app.utils.file_utils import (
    file_name,
    get_default_save_path,
    get_recent_files,
    remember_recent_file,
    set_default_save_path,
    timestamped_file_name,
)
from evaluation_app.utils.validators import name_key


class MainWindow(QMainWindow):
    """Главное окно: список преподавателей, сводная таблица, детали, экспорт."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(config.APP_NAME)
        self.setWindowIcon(app_icon())
        self.resize(1240, 800)

        self.project = Project(project_name="Новый проект", academic_year=default_academic_year())
        self.current_project_path: str | None = None
        self._dirty = False
        self._history: list[dict] = []
        self._sort_key = "score"
        self._sort_reverse = True
        self._search_text = ""
        self._last_directory = get_default_save_path()

        self._create_actions()
        self._create_menus()
        self._create_toolbar()
        self._build_ui()
        self._create_status_bar()

        self._refresh_all()
        self._set_dirty(False)

    # ------------------------------------------------------------------ UI ---

    def _create_actions(self) -> None:
        style = self.style()

        self.action_new = QAction(style.standardIcon(QStyle.SP_FileIcon), "Новый проект...", self)
        self.action_new.setShortcut(QKeySequence.New)
        self.action_new.setStatusTip("Создать новый проект")
        self.action_new.triggered.connect(self._on_new_project)

        self.action_open = QAction(style.standardIcon(QStyle.SP_DialogOpenButton), "Открыть проект...", self)
        self.action_open.setShortcut(QKeySequence.Open)
        self.action_open.triggered.connect(self._on_open_project)

        self.action_save = QAction(style.standardIcon(QStyle.SP_DialogSaveButton), "Сохранить проект", self)
        self.action_save.setShortcut(QKeySequence.Save)
        self.action_save.triggered.connect(self._on_save_project)

        self.action_save_as = QAction("Сохранить как...", self)
        self.action_save_as.setShortcut(QKeySequence("Ctrl+Shift+S"))
        self.action_save_as.triggered.connect(lambda: self._on_save_project(save_as=True))

        self.action_import = QAction(style.standardIcon(QStyle.SP_ArrowDown), "Импорт Excel...", self)
        self.action_import.setShortcut(QKeySequence("Ctrl+I"))
        self.action_import.setStatusTip("Загрузить оценки из файлов преподавателей")
        self.action_import.triggered.connect(self._on_import_excel)

        self.action_template = QAction("Создать шаблон Excel...", self)
        self.action_template.triggered.connect(self._on_create_template)

        self.action_export_summary = QAction(style.standardIcon(QStyle.SP_ArrowUp), "Экспорт сводной таблицы...", self)
        self.action_export_summary.setShortcut(QKeySequence("Ctrl+E"))
        self.action_export_summary.triggered.connect(self._on_export_summary)

        self.action_export_project = QAction("Экспорт проекта в Excel...", self)
        self.action_export_project.triggered.connect(self._on_export_project)

        self.action_add_teacher = QAction(
            style.standardIcon(QStyle.SP_FileDialogNewFolder), "Добавить преподавателя...", self
        )
        self.action_add_teacher.setShortcut(QKeySequence("Ctrl+T"))
        self.action_add_teacher.triggered.connect(self._on_add_teacher)

        self.action_rename_teacher = QAction("Переименовать преподавателя...", self)
        self.action_rename_teacher.triggered.connect(self._on_rename_teacher)

        self.action_remove_teacher = QAction(style.standardIcon(QStyle.SP_TrashIcon), "Удалить преподавателя", self)
        self.action_remove_teacher.triggered.connect(self._on_remove_teacher)

        self.action_update_teacher = QAction("Обновить данные из Excel...", self)
        self.action_update_teacher.triggered.connect(self._on_update_teacher)

        self.action_undo = QAction(style.standardIcon(QStyle.SP_ArrowBack), "Отменить последнее изменение", self)
        self.action_undo.setShortcut(QKeySequence.Undo)
        self.action_undo.triggered.connect(self._on_undo)

        self.action_settings = QAction(style.standardIcon(QStyle.SP_FileDialogDetailedView), "Настройки...", self)
        self.action_settings.triggered.connect(self._on_settings)

        self.action_about = QAction("О программе", self)
        self.action_about.triggered.connect(self._on_about)

        self.action_quit = QAction("Выход", self)
        self.action_quit.setShortcut(QKeySequence.Quit)
        self.action_quit.triggered.connect(self.close)

    def _create_menus(self) -> None:
        file_menu = self.menuBar().addMenu("&Файл")
        file_menu.addAction(self.action_new)
        file_menu.addAction(self.action_open)
        self.recent_menu = file_menu.addMenu("Недавние проекты")
        file_menu.addSeparator()
        file_menu.addAction(self.action_save)
        file_menu.addAction(self.action_save_as)
        file_menu.addSeparator()
        file_menu.addAction(self.action_import)
        file_menu.addAction(self.action_template)
        file_menu.addSeparator()
        file_menu.addAction(self.action_export_summary)
        file_menu.addAction(self.action_export_project)
        file_menu.addSeparator()
        file_menu.addAction(self.action_quit)

        edit_menu = self.menuBar().addMenu("&Правка")
        edit_menu.addAction(self.action_undo)
        edit_menu.addSeparator()
        edit_menu.addAction(self.action_add_teacher)
        edit_menu.addAction(self.action_rename_teacher)
        edit_menu.addAction(self.action_update_teacher)
        edit_menu.addAction(self.action_remove_teacher)
        edit_menu.addSeparator()
        edit_menu.addAction(self.action_settings)

        help_menu = self.menuBar().addMenu("&Справка")
        help_menu.addAction(self.action_about)

        self._rebuild_recent_menu()

    def _create_toolbar(self) -> None:
        toolbar = QToolBar("Основные действия", self)
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        toolbar.addAction(self.action_new)
        toolbar.addAction(self.action_open)
        toolbar.addAction(self.action_save)
        toolbar.addSeparator()
        toolbar.addAction(self.action_import)
        toolbar.addAction(self.action_export_summary)
        toolbar.addSeparator()
        toolbar.addAction(self.action_add_teacher)
        toolbar.addAction(self.action_update_teacher)
        toolbar.addAction(self.action_remove_teacher)
        toolbar.addSeparator()
        toolbar.addAction(self.action_undo)
        toolbar.addAction(self.action_settings)

    def _build_ui(self) -> None:
        self.teacher_list = TeacherListWidget(self)
        self.summary_view = SummaryTableView(self)
        self.direction_view = DirectionView(self)

        right_splitter = QSplitter(Qt.Vertical, self)
        right_splitter.addWidget(self.summary_view)
        right_splitter.addWidget(self.direction_view)
        right_splitter.setStretchFactor(0, 3)
        right_splitter.setStretchFactor(1, 2)

        splitter = QSplitter(Qt.Horizontal, self)
        splitter.addWidget(self.teacher_list)
        splitter.addWidget(right_splitter)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([320, 900])

        container = QWidget(self)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.addWidget(splitter)
        self.setCentralWidget(container)

        # Сигналы
        self.teacher_list.addRequested.connect(self._on_add_teacher)
        self.teacher_list.renameRequested.connect(self._on_rename_teacher)
        self.teacher_list.removeRequested.connect(self._on_remove_teacher)
        self.teacher_list.importRequested.connect(self._on_update_teacher)
        self.teacher_list.teacherSelected.connect(self._on_teacher_selected)

        self.summary_view.searchChanged.connect(self._on_search_changed)
        self.summary_view.sortChanged.connect(self._on_sort_changed)
        self.summary_view.exportRequested.connect(self._on_export_summary)
        self.summary_view.sort_widget.set_sort(self._sort_key, self._sort_reverse)

        self.direction_view.exportRequested.connect(self._on_export_teacher)

    def _create_status_bar(self) -> None:
        self.status_label = QLabel("", self)
        self.statusBar().addPermanentWidget(self.status_label)
        self.statusBar().showMessage("Готово")

    # ------------------------------------------------------------ Обновление ---

    def _refresh_all(self, selected_teacher_id: str = "") -> None:
        """Перерисовывает все представления."""
        self.teacher_list.set_teachers(self.project.teachers, selected_teacher_id or None)
        self._refresh_summary()
        self._refresh_direction_view()
        self._update_title()

    def _refresh_summary(self) -> None:
        all_records = DataMerger.merge_by_student(self.project)
        query = name_key(self._search_text)
        records = all_records
        if query:
            records = [r for r in all_records if query in name_key(r.get("student_name", ""))]
        records = DataSorter.sort(records, key=self._sort_key, reverse=self._sort_reverse)

        self.summary_view.set_records(records, total_count=len(all_records))

        # Сводная уже посчитана выше: статистика не пересчитывает её повторно.
        stats = DataMerger.build_statistics(self.project, records=all_records)
        self.status_label.setText(DataMerger.format_statistics(stats))

    def _refresh_direction_view(self) -> None:
        teacher_id = self.teacher_list.current_teacher_id()
        teacher = self.project.get_teacher(teacher_id) if teacher_id else None
        self.direction_view.show_teacher(teacher)

    def _update_title(self) -> None:
        marker = "*" if self._dirty else ""
        name = self.project.project_name or "Без названия"
        if self.project.academic_year:
            name = f"{name} ({self.project.academic_year})"
        self.setWindowTitle(f"{name}{marker} — {config.APP_NAME}")
        self.setWindowFilePath(self.current_project_path or "")

    def _set_dirty(self, value: bool) -> None:
        self._dirty = value
        self._update_title()

    # -------------------------------------------------------------- Слоты ---

    def _on_new_project(self) -> None:
        if not self._confirm_discard():
            return

        dialog = ProjectSettingsDialog(self, project_name="", academic_year=default_academic_year())
        if dialog.exec_() != QDialog.Accepted:
            return

        self._push_history()
        self.project = Project(
            project_name=dialog.get_project_name(),
            academic_year=dialog.get_academic_year(),
        )
        self.current_project_path = None
        self._search_text = ""
        self.summary_view.set_search_text("")
        self._set_dirty(False)
        self._refresh_all()
        self.statusBar().showMessage(f"Создан проект «{self.project.project_name}»", 5000)

    def _on_open_project(self, file_path: str | None = None) -> None:
        if not self._confirm_discard():
            return

        if not file_path:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Открыть проект", self._last_directory, config.PROJECT_FILTER
            )
            if not file_path:
                return

        self._last_directory = os.path.dirname(os.path.abspath(file_path))
        try:
            project = ProjectManager.load_project(file_path)
        except ProjectError as exc:
            QMessageBox.critical(self, "Ошибка открытия проекта", str(exc))
            return

        self.project = project
        self.current_project_path = file_path
        self._history.clear()
        remember_recent_file(file_path)
        self._rebuild_recent_menu()
        self._search_text = ""
        self.summary_view.set_search_text("")
        self._set_dirty(False)
        self._refresh_all()
        self.statusBar().showMessage(
            f"Проект загружен: {file_name(file_path)} "
            f"({len(self.project.teachers)} преподавателей)", 6000
        )

    def _on_save_project(self, save_as: bool = False) -> None:
        file_path = self.current_project_path if not save_as else None
        if not file_path:
            default_name = (self.project.project_name or "project").strip() or "project"
            suggested = os.path.join(
                get_default_save_path(), f"{default_name}{config.PROJECT_EXTENSION}"
            )
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Сохранить проект", suggested, config.PROJECT_FILTER
            )
            if not file_path:
                return
            if not file_path.lower().endswith(config.PROJECT_EXTENSION):
                file_path += config.PROJECT_EXTENSION

        try:
            ProjectManager.save_project(self.project, file_path)
        except ProjectError as exc:
            QMessageBox.critical(self, "Ошибка сохранения", str(exc))
            return

        self.current_project_path = file_path
        self._last_directory = os.path.dirname(os.path.abspath(file_path))
        remember_recent_file(file_path)
        self._rebuild_recent_menu()
        self._set_dirty(False)
        self.statusBar().showMessage(f"Проект сохранён: {file_name(file_path)}", 5000)

    def _on_import_excel(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Выберите Excel-файлы", self._last_directory, config.EXCEL_FILTER
        )
        if not paths:
            return
        self._last_directory = os.path.dirname(os.path.abspath(paths[0]))

        # Разбор в фоновом потоке: окно не подвисает, виден прогресс.
        outcome = run_parse(paths, self)
        if outcome is None:
            self.statusBar().showMessage("Импорт отменён", 4000)
            return
        if outcome.errors:
            QMessageBox.critical(self, "Ошибка чтения файла", "\n\n".join(outcome.errors))

        items = outcome.items
        if not items:
            QMessageBox.warning(self, "Импорт", "В выбранных файлах не найдено данных для импорта.")
            return

        dialog = UploadPreviewDialog(
            items, [t.name for t in self.project.teachers], self
        )
        dialog.configureRequested.connect(lambda row: self._on_configure_columns(dialog, row))
        if dialog.exec_() != QDialog.Accepted:
            for parser in {item.parser for item in items if item.parser is not None}:
                parser.close()
            return

        selection = dialog.get_selection()
        if not selection:
            QMessageBox.information(self, "Импорт", "Все листы отмечены как «пропустить».")
            for parser in {item.parser for item in items if item.parser is not None}:
                parser.close()
            return

        self._push_history()
        added, replaced, warnings = self.apply_import_sheets(selection)
        for parser in {item.parser for item in items if item.parser is not None}:
            parser.close()
        self._set_dirty(True)
        self._refresh_all()

        message = (
            f"Добавлено преподавателей: {added}\n"
            f"Обновлено направлений: {replaced}\n"
            f"Всего оценок в проекте: {len(self.project.evaluations())}"
        )
        if warnings:
            message += f"\n\nПредупреждений: {len(warnings)}"
            self._show_warnings(warnings, title="Импорт завершён с предупреждениями")
        else:
            QMessageBox.information(self, "Импорт завершён", message)
        self.statusBar().showMessage(
            f"Импорт из Excel: +{added} преподавателей, обновлено направлений: {replaced}", 7000
        )

    def _collect_upload_items(self, paths: Sequence[str]) -> list[UploadSheet]:
        """Читает файлы и готовит по каждому листу элемент предпросмотра.

        Используется только в тестах (синхронный запасной путь): основной
        сценарий импорта разбирает файлы в фоне через :func:`run_parse`.
        """
        outcome = collect_upload_items(paths)
        if outcome.errors:
            QMessageBox.critical(self, "Ошибка чтения файла", "\n\n".join(outcome.errors))
        return outcome.items

    def _on_configure_columns(self, dialog: UploadPreviewDialog, row: int) -> None:
        """Ручное определение столбцов для листа из диалога импорта."""
        item = dialog.item_at(row)
        if item is None or item.parser is None:
            return

        mapping_dialog = ColumnMappingDialog(item.parser, item.sheet_name, self)
        if mapping_dialog.exec_() != QDialog.Accepted:
            return

        header_row = mapping_dialog.header_row()
        columns = mapping_dialog.columns()
        mode = mapping_dialog.mode()
        try:
            item.result = item.parser.parse_sheet(
                item.sheet_name, header_row=header_row, columns=columns, mode=mode
            )
            item.error = None
            self.statusBar().showMessage(
                f"Структура листа «{item.sheet_name}» настроена: "
                f"режим {item.result.mode}, оценок {item.result.evaluations_count}",
                6000,
            )
        except ExcelParseError as exc:
            item.error = str(exc)
            QMessageBox.warning(self, "Не удалось разобрать лист", str(exc))
        dialog.refresh()

    def _apply_import_selection(self, selection: Sequence[tuple]) -> tuple[int, int, list[str]]:
        """Применяет выбор диалога импорта к проекту (устаревший формат вызова).

        Используется только в тестах: принимает тройки «(файл, лист, ФИО)» и
        разбирает файлы заново. Приложение вызывает :meth:`apply_import_sheets`
        с уже разобранными листами из диалога предпросмотра — так учитываются
        ручные настройки столбцов и файлы не читаются второй раз.
        """
        sheets: list[tuple[UploadSheet, str | None]] = []
        parse_errors: list[str] = []
        for file_path, sheet_name, teacher_name in selection:
            parser = ExcelParser(file_path)
            try:
                parser.load()
                result = parser.parse_sheet(sheet_name)
            except ExcelParseError as exc:
                parse_errors.append(f"{file_name(file_path)} / {sheet_name}: {exc}")
                parser.close()
                continue
            parser.close()
            sheets.append(
                (
                    UploadSheet(
                        file_path=file_path,
                        sheet_name=sheet_name,
                        parser=parser,
                        result=result,
                    ),
                    teacher_name,
                )
            )
        added, replaced, warnings = self.apply_import_sheets(sheets)
        return added, replaced, [*parse_errors, *warnings]

    def apply_import_sheets(
        self, sheets: Sequence[tuple[UploadSheet, str | None]]
    ) -> tuple[int, int, list[str]]:
        """Добавляет в проект уже разобранные листы из диалога импорта.

        Файлы заново не читаются: результаты разбора, включая ручные настройки
        столбцов и режима, уже лежат в ``item.result``.

        :param sheets: пары «элемент предпросмотра → ФИО преподавателя».
        :return: (добавлено преподавателей, обновлено направлений, предупреждения).
        """
        added = 0
        replaced = 0
        warnings: list[str] = []

        for item, teacher_name in sheets:
            result = item.result
            if result is None:
                continue

            teacher_name = (teacher_name or "").strip()
            teacher = self.project.find_teacher(teacher_name)
            if teacher is None:
                teacher = self.project.add_teacher(name=teacher_name)
                added += 1

            existing = teacher.get_direction(result.direction.name)
            if existing is not None:
                existing.replace_evaluations(result.direction.evaluations)
                replaced += 1
            else:
                teacher.directions.append(result.direction)

            warnings.extend(
                f"{file_name(item.file_path)} / {item.sheet_name}: {warning}"
                for warning in result.warnings
            )

        return added, replaced, warnings

    def _on_update_teacher(self, teacher_id: str = "") -> None:
        """Обновление данных преподавателя из Excel (старые данные заменяются)."""
        teacher_id = teacher_id or self.teacher_list.current_teacher_id()
        teacher = self.project.get_teacher(teacher_id)
        if teacher is None:
            QMessageBox.warning(self, "Обновление", "Выберите преподавателя в списке слева.")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            f"Выберите файл с данными для «{teacher.name}»",
            self._last_directory,
            config.EXCEL_FILTER,
        )
        if not file_path:
            return
        self._last_directory = os.path.dirname(os.path.abspath(file_path))

        parser = ExcelParser(file_path, teacher_name=teacher.name)
        try:
            parser.load()
            sheet_names = parser.sheet_names
        except ExcelParseError as exc:
            QMessageBox.critical(self, "Ошибка чтения файла", str(exc))
            return

        directions: list[Direction] = []
        warnings: list[str] = []
        for sheet_name in sheet_names:
            try:
                result = parser.parse_sheet(sheet_name)
            except ExcelParseError as exc:
                # Предлагаем вручную указать столбцы и режим разбора.
                mapping_dialog = ColumnMappingDialog(parser, sheet_name, self)
                if mapping_dialog.exec_() != QDialog.Accepted:
                    warnings.append(f"{sheet_name}: {exc}")
                    continue
                try:
                    result = parser.parse_sheet(
                        sheet_name,
                        header_row=mapping_dialog.header_row(),
                        columns=mapping_dialog.columns(),
                        mode=mapping_dialog.mode(),
                    )
                except ExcelParseError as inner:
                    warnings.append(f"{sheet_name}: {inner}")
                    continue
            directions.append(result.direction)
            warnings.extend(f"{sheet_name}: {w}" for w in result.warnings)

        if not directions:
            QMessageBox.warning(
                self, "Обновление", "Не удалось получить данные ни с одного листа файла."
            )
            return

        answer = QMessageBox.question(
            self,
            "Подтвердите обновление",
            f"Заменить данные преподавателя «{teacher.name}»?\n\n"
            f"Новых направлений: {len(directions)}\n"
            f"Оценок: {sum(len(d.evaluations) for d in directions)}\n\n"
            "Старые данные будут удалены.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        parser.close()
        if answer != QMessageBox.Yes:
            return

        self._push_history()
        ProjectManager.replace_teacher_directions(self.project, teacher.id, directions)
        self._set_dirty(True)
        self._refresh_all(selected_teacher_id=teacher.id)

        if warnings:
            self._show_warnings(warnings, title="Данные обновлены с предупреждениями")
        else:
            QMessageBox.information(
                self, "Обновление", f"Данные преподавателя «{teacher.name}» обновлены."
            )

    def _on_add_teacher(self) -> None:
        name, ok = QInputDialog.getText(self, "Новый преподаватель", "ФИО преподавателя:")
        if not ok or not name.strip():
            return

        name = name.strip()
        if self.project.find_teacher(name) is not None:
            QMessageBox.warning(self, "Преподаватель", f"Преподаватель «{name}» уже есть в проекте.")
            return

        self._push_history()
        teacher = self.project.add_teacher(name=name)
        self._set_dirty(True)
        self._refresh_all(selected_teacher_id=teacher.id)
        self.statusBar().showMessage(f"Добавлен преподаватель: {name}", 5000)

    def _on_rename_teacher(self, teacher_id: str = "") -> None:
        teacher_id = teacher_id or self.teacher_list.current_teacher_id()
        teacher = self.project.get_teacher(teacher_id)
        if teacher is None:
            QMessageBox.warning(self, "Переименование", "Выберите преподавателя в списке слева.")
            return

        name, ok = QInputDialog.getText(
            self, "Переименование", "Новое ФИО преподавателя:", text=teacher.name
        )
        if not ok or not name.strip() or name.strip() == teacher.name:
            return

        self._push_history()
        self.project.rename_teacher(teacher.id, name.strip())
        self._set_dirty(True)
        self._refresh_all(selected_teacher_id=teacher.id)

    def _on_remove_teacher(self, teacher_id: str = "") -> None:
        teacher_id = teacher_id or self.teacher_list.current_teacher_id()
        teacher = self.project.get_teacher(teacher_id)
        if teacher is None:
            QMessageBox.warning(self, "Удаление", "Выберите преподавателя в списке слева.")
            return

        answer = QMessageBox.question(
            self,
            "Удаление преподавателя",
            f"Удалить преподавателя «{teacher.name}»?\n\n"
            f"Будут удалены направления: {len(teacher.directions)}, "
            f"оценки: {len(teacher.evaluations())}.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return

        self._push_history()
        self.project.remove_teacher(teacher.id)
        self._set_dirty(True)
        self._refresh_all()
        self.statusBar().showMessage(f"Удалён преподаватель: {teacher.name}", 5000)

    def _on_search_changed(self, text: str) -> None:
        self._search_text = text
        self._refresh_summary()

    def _on_sort_changed(self, key: str, reverse: bool) -> None:
        self._sort_key = key
        self._sort_reverse = reverse
        self._refresh_summary()

    def _on_teacher_selected(self, teacher_id: str) -> None:
        teacher = self.project.get_teacher(teacher_id) if teacher_id else None
        self.direction_view.show_teacher(teacher)

    def _on_undo(self) -> None:
        if not self._history:
            self.statusBar().showMessage("Нет изменений для отмены", 4000)
            return

        snapshot = self._history.pop()
        self.project = Project.from_dict(snapshot)
        self._set_dirty(True)
        self._refresh_all()
        self.statusBar().showMessage(
            f"Изменение отменено (доступно отмен: {len(self._history)})", 5000
        )

    def _on_settings(self) -> None:
        dialog = AppSettingsDialog(self, default_save_path=get_default_save_path())
        if dialog.exec_() != QDialog.Accepted:
            return
        path = dialog.get_default_save_path()
        if path:
            set_default_save_path(path)
            self._last_directory = path
            self.statusBar().showMessage(f"Каталог сохранения: {path}", 5000)

    def _on_about(self) -> None:
        QMessageBox.about(
            self,
            "О программе",
            f"<b>{config.APP_NAME}</b> {config.APP_VERSION}<br><br>"
            "Сбор и обработка оценок преподавателей.<br>"
            "Загрузка Excel-файлов, объединение баллов по ФИО учеников, "
            "сортировка, экспорт результатов и сохранение проекта в JSON.<br><br>"
            f"Python + PyQt5, openpyxl<br>Каталог по умолчанию: {get_default_save_path()}",
        )

    # --- Экспорт -------------------------------------------------------------

    def _on_export_summary(self) -> None:
        records = self._visible_records()
        if not records:
            QMessageBox.warning(self, "Экспорт", "В сводной таблице нет данных для экспорта.")
            return

        prefix = (self.project.project_name or "Сводная таблица").strip() or "Сводная таблица"
        suggested = os.path.join(get_default_save_path(), timestamped_file_name(prefix, ".xlsx"))
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Экспорт сводной таблицы", suggested, config.EXCEL_FILTER
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            ExcelExporter.export_summary(records, file_path, project=self.project)
        except OSError as exc:
            QMessageBox.critical(self, "Ошибка экспорта", str(exc))
            return
        self.statusBar().showMessage(f"Сводная таблица экспортирована: {file_name(file_path)}", 7000)

    def _on_export_project(self) -> None:
        if self.project.is_empty():
            QMessageBox.warning(self, "Экспорт", "Проект пуст — нечего экспортировать.")
            return

        prefix = (self.project.project_name or "Проект").strip() or "Проект"
        suggested = os.path.join(
            get_default_save_path(), timestamped_file_name(f"{prefix}_полный", ".xlsx")
        )
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Экспорт проекта", suggested, config.EXCEL_FILTER
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            ExcelExporter.export_project(self.project, file_path)
        except OSError as exc:
            QMessageBox.critical(self, "Ошибка экспорта", str(exc))
            return
        self.statusBar().showMessage(f"Проект экспортирован: {file_name(file_path)}", 7000)

    def _on_export_teacher(self, teacher_id: str) -> None:
        teacher = self.project.get_teacher(teacher_id)
        if teacher is None or not teacher.directions:
            QMessageBox.warning(self, "Экспорт", "У преподавателя нет данных для экспорта.")
            return

        suggested = os.path.join(
            get_default_save_path(), timestamped_file_name(teacher.name or "Преподаватель", ".xlsx")
        )
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Экспорт данных преподавателя", suggested, config.EXCEL_FILTER
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            temp_project = Project(project_name=teacher.name, academic_year=self.project.academic_year)
            temp_project.teachers.append(teacher)
            ExcelExporter.export_project(temp_project, file_path)
        except OSError as exc:
            QMessageBox.critical(self, "Ошибка экспорта", str(exc))
            return
        self.statusBar().showMessage(f"Данные преподавателя экспортированы: {file_name(file_path)}", 7000)

    def _on_create_template(self) -> None:
        suggested = os.path.join(
            get_default_save_path(), timestamped_file_name("Шаблон_оценок", ".xlsx")
        )
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Создать шаблон Excel", suggested, config.EXCEL_FILTER
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            ExcelExporter.create_template(file_path)
        except OSError as exc:
            QMessageBox.critical(self, "Ошибка", str(exc))
            return

        QMessageBox.information(
            self,
            "Шаблон создан",
            f"Шаблон сохранён:\n{file_path}\n\n"
            "Передайте файл преподавателям: они заполняют листы (каждое направление — "
            "отдельный лист) и возвращают файл для загрузки в программу.",
        )

    # --- Вспомогательные методы ---------------------------------------------

    def _visible_records(self) -> list[dict]:
        """Записи сводной таблицы с учётом поиска и сортировки."""
        records = DataMerger.merge_by_student(self.project)
        query = name_key(self._search_text)
        if query:
            records = [r for r in records if query in name_key(r.get("student_name", ""))]
        return DataSorter.sort(records, key=self._sort_key, reverse=self._sort_reverse)

    def _push_history(self) -> None:
        """Сохраняет снимок проекта для отмены изменения."""
        self._history.append(self.project.to_dict())
        if len(self._history) > config.MAX_HISTORY:
            self._history.pop(0)
        self.action_undo.setEnabled(True)

    def _confirm_discard(self) -> bool:
        """Спрашивает подтверждение, если есть несохранённые изменения."""
        if not self._dirty:
            return True
        answer = QMessageBox.question(
            self,
            "Несохранённые изменения",
            "В проекте есть несохранённые изменения. Сохранить их?",
            QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
            QMessageBox.Save,
        )
        if answer == QMessageBox.Save:
            self._on_save_project()
            return not self._dirty
        return answer == QMessageBox.Discard

    def _show_warnings(self, warnings: Sequence[str], title: str = "Предупреждения") -> None:
        """Показывает список предупреждений (не более 30 строк)."""
        shown = list(warnings)[:30]
        text = "\n".join(f"• {warning}" for warning in shown)
        if len(warnings) > len(shown):
            text += f"\n...и ещё {len(warnings) - len(shown)}"
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setIcon(QMessageBox.Warning)
        box.setText(
            f"Обнаружены проблемы ({len(warnings)}). "
            "Подробности — в разделе «Показать подробности»."
        )
        box.setDetailedText(text)
        box.exec_()

    def _rebuild_recent_menu(self) -> None:
        """Обновляет подменю недавних проектов."""
        self.recent_menu.clear()
        recent = get_recent_files()
        if not recent:
            empty = self.recent_menu.addAction("Список пуст")
            empty.setEnabled(False)
            return
        for path in recent:
            action = self.recent_menu.addAction(file_name(path))
            action.setToolTip(path)
            action.triggered.connect(lambda _checked=False, p=path: self._on_open_project(p))

    def closeEvent(self, event) -> None:
        """Подтверждение выхода при несохранённых изменениях."""
        if self._confirm_discard():
            event.accept()
        else:
            event.ignore()
