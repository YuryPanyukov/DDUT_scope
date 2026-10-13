"""Экспорт данных и создание шаблонов Excel (см. п. 8.3 плана)."""

from __future__ import annotations

import os
from collections.abc import Iterable, Sequence

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from evaluation_app.models import Direction, Evaluation, Project
from evaluation_app.services.data_merger import DataMerger
from evaluation_app.services.data_sorter import DataSorter

HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
HEADER_FONT = Font(bold=True)
TITLE_FONT = Font(bold=True, size=13)
THIN_BORDER = Border(
    left=Side(style="thin", color="B7B7B7"),
    right=Side(style="thin", color="B7B7B7"),
    top=Side(style="thin", color="B7B7B7"),
    bottom=Side(style="thin", color="B7B7B7"),
)

TEMPLATE_HEADERS = ("ФИО", "Название конкурса", "Итого")
TEMPLATE_ROWS: tuple[tuple[str, str, float], ...] = (
    ("Иванов Иван Иванович", "Олимпиада по математике", 85),
    ("Петрова Анна Сергеевна", "Олимпиада по математике", 92),
    ("Сидоров Пётр Алексеевич", "Конкурс исследовательских работ", 78),
)
TEMPLATE_DIRECTIONS = ("Направление 1", "Направление 2")


class ExcelExporter:
    """Запись данных в Excel: шаблоны, сводные таблицы, экспорт проекта."""

    # --- Вспомогательные методы ---------------------------------------------

    @staticmethod
    def _style_header_row(worksheet, row: int, columns: int) -> None:
        for column in range(1, columns + 1):
            cell = worksheet.cell(row=row, column=column)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = THIN_BORDER

    @staticmethod
    def _autosize(worksheet, widths: Sequence[int]) -> None:
        for index, width in enumerate(widths, start=1):
            worksheet.column_dimensions[get_column_letter(index)].width = width

    @staticmethod
    def _finish(worksheet, header_row: int, columns: int, widths: Sequence[int]) -> None:
        ExcelExporter._style_header_row(worksheet, header_row, columns)
        ExcelExporter._autosize(worksheet, widths)
        worksheet.freeze_panes = f"A{header_row + 1}"
        if worksheet.max_row > header_row:
            worksheet.auto_filter.ref = (
                f"A{header_row}:{get_column_letter(columns)}{worksheet.max_row}"
            )

    @staticmethod
    def _save(workbook, file_path: str) -> str:
        directory = os.path.dirname(os.path.abspath(file_path))
        if directory:
            os.makedirs(directory, exist_ok=True)
        try:
            workbook.save(file_path)
        except OSError as exc:
            raise OSError(f"Не удалось сохранить файл «{file_path}»: {exc}") from exc
        return file_path

    # --- Шаблон для преподавателей ------------------------------------------

    @staticmethod
    def create_template(
        file_path: str,
        sheet_names: Iterable[str] = TEMPLATE_DIRECTIONS,
        rows: Iterable[Sequence] | None = None,
    ) -> str:
        """Создаёт шаблон Excel для заполнения преподавателем.

        На каждом листе направления записывается строка заголовков и примеры
        строк (серым курсивом) — их преподаватель заменяет своими данными.
        """
        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)

        example_rows = list(rows) if rows is not None else list(TEMPLATE_ROWS)
        names = [str(name) for name in sheet_names] or ["Направление 1"]

        for sheet_name in names:
            worksheet = workbook.create_sheet(title=sheet_name[:31])
            worksheet.append(list(TEMPLATE_HEADERS))
            for number, row in enumerate(example_rows, start=1):
                worksheet.append([row[0], row[1], row[2]])
                # Примеры заливаем бледным фоном — их нужно заменить своими данными.
                for column in range(1, 4):
                    cell = worksheet.cell(row=number + 1, column=column)
                    cell.font = Font(italic=True, color="808080")
            ExcelExporter._finish(worksheet, header_row=1, columns=3, widths=(32, 40, 12))
            for row_idx in range(2, worksheet.max_row + 1):
                worksheet.cell(row=row_idx, column=3).number_format = "0.##"

        return ExcelExporter._save(workbook, file_path)

    # --- Экспорт результатов -------------------------------------------------

    @staticmethod
    def export_summary(
        records: Iterable[dict],
        file_path: str,
        project: Project | None = None,
        sheet_title: str = "Сводная таблица",
    ) -> str:
        """Экспортирует сводную таблицу (ФИО → итого баллов)."""
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = sheet_title[:31]

        headers = (
            "ФИО",
            "Итого баллов",
            "Оценок",
            "Преподаватели",
            "Направления",
            "Конкурсы",
        )

        # Заголовок отчёта пишем напрямую в ячейки: append() использует
        # внутренний счётчик строк и перезаписал бы строку с заголовком.
        header_row = 1
        if project is not None:
            title = project.project_name or "Оценочный комитет"
            if project.academic_year:
                title = f"{title} ({project.academic_year})"
            worksheet.cell(row=1, column=1, value=title).font = TITLE_FONT
            header_row = 3

        for column, header in enumerate(headers, start=1):
            worksheet.cell(row=header_row, column=column, value=header)

        for offset, record in enumerate(records, start=1):
            values = (
                record.get("student_name", ""),
                float(record.get("total_score") or 0.0),
                int(record.get("entries") or 0),
                ", ".join(record.get("teachers_involved") or []),
                ", ".join(record.get("directions") or []),
                ", ".join(record.get("competitions") or []),
            )
            for column, value in enumerate(values, start=1):
                worksheet.cell(row=header_row + offset, column=column, value=value)

        ExcelExporter._finish(
            worksheet,
            header_row=header_row,
            columns=6,
            widths=(32, 14, 10, 28, 28, 46),
        )
        for row_idx in range(header_row + 1, worksheet.max_row + 1):
            worksheet.cell(row=row_idx, column=2).number_format = "0.##"

        return ExcelExporter._save(workbook, file_path)

    @staticmethod
    def export_project(project: Project, file_path: str) -> str:
        """Экспортирует весь проект: лист «Сводная» + листы по преподавателям."""
        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)

        summary = workbook.create_sheet(title="Сводная")
        summary.append(["ФИО", "Итого баллов", "Оценок", "Преподаватели", "Направления", "Конкурсы"])

        for record in DataSorter.sort_by_score(DataMerger.merge_by_student(project), reverse=True):
            summary.append(
                [
                    record.get("student_name", ""),
                    float(record.get("total_score") or 0.0),
                    int(record.get("entries") or 0),
                    ", ".join(record.get("teachers_involved") or []),
                    ", ".join(record.get("directions") or []),
                    ", ".join(record.get("competitions") or []),
                ]
            )
        ExcelExporter._finish(summary, header_row=1, columns=6, widths=(32, 14, 10, 28, 28, 46))
        # Формат баллов — как в export_summary: без «хвостовых» нулей (85, а не 85.0).
        for row_idx in range(2, summary.max_row + 1):
            summary.cell(row=row_idx, column=2).number_format = "0.##"

        used_titles = {"Сводная"}
        for teacher in project.teachers:
            worksheet = workbook.create_sheet(title=ExcelExporter._unique_sheet_title(teacher.name, used_titles))
            worksheet.append([f"Преподаватель: {teacher.name}"])
            worksheet.cell(row=1, column=1).font = TITLE_FONT
            for direction in teacher.directions:
                worksheet.append([f"Направление: {direction.name}"])
                worksheet.cell(row=worksheet.max_row, column=1).font = HEADER_FONT
                worksheet.append(["ФИО", "Название конкурса", "Итого"])
                ExcelExporter._style_header_row(worksheet, worksheet.max_row, 3)
                for student_name, evaluations in DataMerger.group_by_student(direction.evaluations):
                    for evaluation in evaluations:
                        worksheet.append(
                            [
                                student_name,
                                evaluation.competition_name,
                                float(evaluation.total_score),
                            ]
                        )
                    # Итоговая сумма баллов ученика — как в его строке ИТОГО.
                    student_total = float(sum(evaluation.total_score or 0.0 for evaluation in evaluations))
                    worksheet.append(["", f"Итого по ученику: {student_name}", student_total])
                    worksheet.cell(row=worksheet.max_row, column=2).font = HEADER_FONT
                    worksheet.cell(row=worksheet.max_row, column=3).font = HEADER_FONT
                worksheet.append(["", "Итого по направлению", direction.total_score()])
                worksheet.cell(row=worksheet.max_row, column=2).font = HEADER_FONT
                worksheet.cell(row=worksheet.max_row, column=3).font = HEADER_FONT
                worksheet.append([])
            ExcelExporter._autosize(worksheet, (32, 40, 12))
            # Баллы и «Итого по направлению» — тот же формат, что в шаблоне и
            # сводной. В колонке есть и текстовые строки (заголовки направлений),
            # поэтому формат ставится только числовым ячейкам.
            for row_idx in range(1, worksheet.max_row + 1):
                cell = worksheet.cell(row=row_idx, column=3)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "0.##"

        return ExcelExporter._save(workbook, file_path)

    @staticmethod
    def export_direction(direction: Direction, file_path: str) -> str:
        """Экспортирует одно направление в отдельный файл.

        Используется только в тестах: из интерфейса выгружается весь проект
        (:meth:`export_project`).
        """
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = (direction.name or "Направление")[:31]
        worksheet.append(["ФИО", "Название конкурса", "Итого"])
        for evaluation in direction.evaluations:
            worksheet.append(
                [evaluation.student_name, evaluation.competition_name, float(evaluation.total_score)]
            )
        ExcelExporter._finish(worksheet, header_row=1, columns=3, widths=(32, 40, 12))
        return ExcelExporter._save(workbook, file_path)

    @staticmethod
    def _unique_sheet_title(name: str, used: set) -> str:
        """Подбирает уникальное имя листа Excel (максимум 31 символ)."""
        base = "".join(ch for ch in (name or "Преподаватель") if ch not in "[]:*?/\\").strip()
        base = (base or "Преподаватель")[:28]
        title = base
        counter = 2
        while title in used:
            suffix = f" ({counter})"
            title = base[: 31 - len(suffix)] + suffix
            counter += 1
        used.add(title)
        return title

    @staticmethod
    def template_evaluations() -> list[Evaluation]:
        """Примеры оценок шаблона (строки ``TEMPLATE_ROWS``).

        Используется только в тестах: сверяет число строк в созданном шаблоне
        с ожидаемым, не зашивая его в тест повторно.
        """
        return [
            Evaluation(student_name=row[0], competition_name=row[1], total_score=row[2])
            for row in TEMPLATE_ROWS
        ]
