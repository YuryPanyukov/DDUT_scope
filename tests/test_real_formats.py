"""Тесты на реальные форматы оценочных листов.

Формат «критерии» (лист-матрица) взят из настоящих файлов преподавателей:
многоуровневая шапка (группы направлений, названия критериев, максимумы баллов),
служебный столбец «за вход в систему», колонка ИТОГО и тысячи «фантомных» пустых строк.
"""

from __future__ import annotations

import os
import zipfile

import pytest

from evaluation_app.models import Project
from evaluation_app.services.data_merger import DataMerger
from evaluation_app.services.excel_parser import MODE_CRITERIA, MODE_ROWS, ExcelParser
from evaluation_app.utils.file_utils import safe_file_stem
from evaluation_app.utils.validators import is_service_label
from tests.helpers import make_workbook

# Лист-матрица: шапка в 3 строки, служебная колонка 2, критерии в колонках 3–7, ИТОГО в 8.
MATRIX_SHEET = [
    ["ФИО", "за вход в систему", "УЧЕБНАЯ И ПРОЕКТНАЯ ДЕЯТЕЛЬНОСТЬ", None, None,
     "ТВОРЧЕСКАЯ ДЕЯТЕЛЬНОСТЬ", None, "ИТОГО"],
    [None, None, "Участие (внутренний)", "Призовое место (внутреннее)",
     "Призовые места в заочном конкурсе", "Разовое мероприятие",
     "Очное участие в мастер-классах", None],
    [None, 30, 30, "40/50/60", "30/значимые", 30, 50, None],
    ["Соколова Наталья Фёдоровна", 10, 10, 5, 0, 30, 0, 45],
    ["Михайлов Артём Игоревич", 0, 0, 0, 30, 0, 15, 45],
    ["Лебедева Ольга Константиновна", 30, 30, 30, 0, 0, 0, 60],
]

# Тот же лист, но с дополнительной строкой нумерации столбцов (как в файлах учителей).
NUMBERED_SHEET = [
    ["ФИО", 1, 2, 3, 4, 5, 6],
    [None, "за вход в систему", "ПРОЕКТНАЯ ДЕЯТЕЛЬНОСТЬ", None, "СПОРТИВНАЯ ДЕЯТЕЛЬНОСТЬ", None, "ИТОГО"],
    [None, None, "Участие (внутренний)", "Призовое место (внутреннее)",
     "Очное участие в мастер-классах", "Разовое мероприятие", None],
    [None, 30, 30, "40/50/60", 50, 30, None],
    ["Михайлов Артём Игоревич", 10, 10, 0, 50, 0, 60],
    ["Кузнецов Лев Дмитриевич", 0, 0, 30, 0, 30, 60],
]


def test_service_labels_are_recognised() -> None:
    assert is_service_label("за вход в систему")
    assert is_service_label("5 класс")
    assert is_service_label("№")
    assert is_service_label("32")
    assert is_service_label("ФИО педагога")
    # «мастер-классах» — критерий, а не служебный столбец
    assert not is_service_label("Очное участие в открытых мастер-классах")
    assert not is_service_label("Посещаемость 90% занятий")
    assert not is_service_label("Призовые места 3,2,1")


def test_matrix_layout_is_recognised(tmp_path) -> None:
    path = make_workbook(tmp_path / "matrix.xlsx", {"Направление_1": MATRIX_SHEET})
    result = ExcelParser(path).parse_sheet("Направление_1")

    assert result.mode == MODE_CRITERIA
    assert result.header_rows == (1, 3)  # группы, критерии, максимумы
    assert result.columns == {"student_name": 1, "total_score": 8}
    # «за вход в систему» — служебный столбец, ИТОГО — не критерий
    assert result.criterion_columns == (3, 4, 5, 6, 7)
    assert result.criterion_names == (
        "Участие (внутренний)",
        "Призовое место (внутреннее)",
        "Призовые места в заочном конкурсе",
        "Разовое мероприятие",
        "Очное участие в мастер-классах",
    )
    # Формула ИТОГО (=SUM по критериям) не учитывает столбец 2 — это ошибка шаблона,
    # о ней есть предупреждение (у Михайловой в B ноль — предупреждения нет).
    assert any("формула ИТОГО не учитывает столбец" in warning for warning in result.warnings)
    assert any("«за вход в систему»" in warning for warning in result.warnings)


def test_matrix_totals_match_total_column(tmp_path) -> None:
    path = make_workbook(tmp_path / "matrix.xlsx", {"Направление_1": MATRIX_SHEET})
    project = Project(project_name="Тест")
    parser = ExcelParser(path, teacher_name="Орлова А.Б.")
    for teacher in parser.parse_all_teachers():
        project.teachers.append(teacher)

    records = {r["student_name"]: r for r in DataMerger.merge_by_student(project)}
    # Формула ИТОГО пропускает столбец «за вход в систему» — его баллы добавлены к итогу:
    # Соколова 45+10, Михайлова 45+0, Лебедева 60+30.
    assert records["Соколова Наталья Фёдоровна"]["total_score"] == 55
    assert records["Михайлов Артём Игоревич"]["total_score"] == 45
    assert records["Лебедева Ольга Константиновна"]["total_score"] == 90
    assert records["Лебедева Ольга Константиновна"]["competitions"] == [
        "за вход в систему",
        "Участие (внутренний)",
        "Призовое место (внутреннее)",
    ]


def test_matrix_with_numbering_row(tmp_path) -> None:
    path = make_workbook(tmp_path / "numbered.xlsx", {"Направление_2": NUMBERED_SHEET})
    result = ExcelParser(path).parse_sheet("Направление_2")

    assert result.mode == MODE_CRITERIA
    assert result.header_rows == (1, 4)  # нумерация, группа, критерии, максимумы
    assert result.columns == {"student_name": 1, "total_score": 7}
    assert result.criterion_columns == (3, 4, 5, 6)

    direction = result.direction
    totals = {e.student_name: 0.0 for e in direction.evaluations}
    for evaluation in direction.evaluations:
        totals[evaluation.student_name] += evaluation.total_score
    # У Михайлова в B 10 баллов, которых нет в формуле ИТОГО — они прибавлены к итогу.
    assert totals == {"Михайлов Артём Игоревич": 70, "Кузнецов Лев Дмитриевич": 60}


def test_matrix_skips_totals_row(tmp_path) -> None:
    rows = [*MATRIX_SHEET, ["ИТОГО", 40, 40, 35, 30, 30, 15, 150]]
    path = make_workbook(tmp_path / "totals.xlsx", {"Лист": rows})
    result = ExcelParser(path).parse_sheet("Лист")

    students = {e.student_name for e in result.direction.evaluations}
    assert "ИТОГО" not in students
    assert len(students) == 3


def test_matrix_uses_file_total_when_criteria_differ(tmp_path) -> None:
    """Если сумма критериев расходится с колонкой ИТОГО — берём значение из файла."""
    rows = [row[:] for row in MATRIX_SHEET]
    rows[3] = ["Соколова Наталья Фёдоровна", 10, 10, 5, 0, 30, 0, 95]  # ИТОГО «не сходится»

    path = make_workbook(tmp_path / "mismatch.xlsx", {"Лист": rows})
    result = ExcelParser(path).parse_sheet("Лист")

    evaluations = [e for e in result.direction.evaluations if e.student_name.startswith("Соколова")]
    assert len(evaluations) == 1
    assert evaluations[0].total_score == 95
    assert evaluations[0].competition_name == "Итого по листу (из файла)"
    assert any("не совпадает" in warning for warning in result.warnings)


def test_matrix_uses_criteria_when_total_is_empty(tmp_path) -> None:
    """Незаполненная колонка ИТОГО (0) не должна обнулять баллы ученика."""
    rows = [row[:] for row in MATRIX_SHEET]
    rows[3] = ["Соколова Наталья Фёдоровна", 10, 10, 5, 0, 30, 0, 0]

    path = make_workbook(tmp_path / "zero_total.xlsx", {"Лист": rows})
    result = ExcelParser(path).parse_sheet("Лист")

    total = sum(e.total_score for e in result.direction.evaluations if e.student_name.startswith("Соколова"))
    assert total == 45
    assert any("не заполнена" in warning for warning in result.warnings)


def test_matrix_adds_login_column_skipped_by_formula(tmp_path) -> None:
    """Формула ИТОГО без столбца «за вход в систему»: баллы входа добавлены к итогу."""
    rows = [row[:] for row in MATRIX_SHEET]
    rows[3] = ["Соколова Наталья Фёдоровна", 10, 10, 5, 0, 30, 0, 45]  # ИТОГО = сумме критериев

    path = make_workbook(tmp_path / "login.xlsx", {"Лист": rows})
    result = ExcelParser(path).parse_sheet("Лист")

    evaluations = [e for e in result.direction.evaluations if e.student_name.startswith("Соколова")]
    scores = {e.competition_name: e.total_score for e in evaluations}
    assert scores["за вход в систему"] == 10  # столбец B добавлен отдельной оценкой
    assert sum(scores.values()) == 55  # 45 критериев + 10 баллов входа
    assert any("формула ИТОГО не учитывает столбец" in w for w in result.warnings)


def test_matrix_without_login_column_keeps_totals(tmp_path) -> None:
    """Ноль в столбце входа не создаёт ни лишних оценок, ни предупреждений."""
    path = make_workbook(tmp_path / "matrix.xlsx", {"Направление_1": MATRIX_SHEET})
    result = ExcelParser(path).parse_sheet("Направление_1")

    zero_login = [
        e for e in result.direction.evaluations if e.student_name.startswith("Михайлов")
    ]
    assert all(e.competition_name != "за вход в систему" for e in zero_login)


def test_matrix_keeps_student_without_any_score(tmp_path) -> None:
    """Ученик без единого балла попадает в таблицу с 0 (задача 2)."""
    rows = [
        *MATRIX_SHEET,
        ["Фёдоров Сергей Николаевич", 0, 0, 0, 0, 0, 0, 0],
    ]
    path = make_workbook(tmp_path / "zero.xlsx", {"Лист": rows})
    result = ExcelParser(path).parse_sheet("Лист")

    zero_student = [e for e in result.direction.evaluations if e.student_name.startswith("Фёдоров")]
    assert len(zero_student) == 1
    assert zero_student[0].total_score == 0
    assert result.rows_parsed == 4


def test_matrix_works_without_total_column(tmp_path) -> None:
    rows = [row[:7] for row in MATRIX_SHEET]  # убрали колонку ИТОГО
    path = make_workbook(tmp_path / "no_total.xlsx", {"Лист": rows})
    result = ExcelParser(path).parse_sheet("Лист")

    assert result.mode == MODE_CRITERIA
    assert result.columns == {"student_name": 1}
    total = sum(e.total_score for e in result.direction.evaluations if e.student_name.startswith("Лебедева"))
    assert total == 60


def test_rows_mode_skips_repeated_header_and_total_rows(tmp_path) -> None:
    rows = [
        ["ФИО", "Название конкурса", "Итого"],
        ["Иванов Иван", "Олимпиада", 85],
        ["ФИО", "Название конкурса", "Итого"],  # повтор шапки при печати
        ["Петрова Анна", "Олимпиада", 92],
        ["ИТОГО", "", 177],
    ]
    path = make_workbook(tmp_path / "repeat.xlsx", {"Лист": rows})
    result = ExcelParser(path).parse_sheet("Лист")

    assert [e.student_name for e in result.direction.evaluations] == ["Иванов Иван", "Петрова Анна"]
    assert result.mode == MODE_ROWS


def test_phantom_rows_are_not_scanned_forever(tmp_path) -> None:
    """Листы реальных файлов содержат тысячи пустых строк — разбор должен их пережить."""
    rows = [*MATRIX_SHEET, *[[None] * 8] * 400]
    path = make_workbook(tmp_path / "phantom.xlsx", {"Лист": rows})
    result = ExcelParser(path).parse_sheet("Лист")

    assert len({e.student_name for e in result.direction.evaluations}) == 3


def test_summary_sheet_warning(tmp_path) -> None:
    path = make_workbook(tmp_path / "admin.xlsx", {"Общий свод": MATRIX_SHEET})
    result = ExcelParser(path).parse_sheet("Общий свод")
    assert any("двойной учёт" in warning for warning in result.warnings)


def test_external_links_detection(tmp_path) -> None:
    path = tmp_path / "svod.xlsx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/externalLinks/externalLink1.xml", "<externalLink/>")

    parser = ExcelParser(str(path))
    assert parser._detect_external_links() == ["xl/externalLinks/externalLink1.xml"]


def test_manual_criteria_mode(tmp_path) -> None:
    """Ручной режим «лист-матрица»: указан только столбец с ФИО."""
    path = make_workbook(tmp_path / "manual.xlsx", {"Лист": MATRIX_SHEET})
    parser = ExcelParser(path)
    parser.load()
    result = parser.parse_sheet("Лист", header_row=1, columns={"student_name": 1}, mode=MODE_CRITERIA)

    assert result.mode == MODE_CRITERIA
    assert result.columns["student_name"] == 1
    assert result.criterion_columns == (3, 4, 5, 6, 7)
    assert result.evaluations_count >= 3


def test_safe_file_stem_builds_teacher_name() -> None:
    assert safe_file_stem("Оценки_Петрова_2025.xlsx") == "Петрова"
    assert safe_file_stem("Админ.xlsx") == "Админ"
    assert safe_file_stem("Критерии_Орлова_АБ.xlsx") == "Орлова А.Б."
    assert safe_file_stem("Иванов ПП 2024-2025.xlsx") == "Иванов П.П."
    # Из одних служебных слов ФИО не собрать — остаётся имя файла
    assert safe_file_stem("Шаблон_оценок.xlsx") == "Шаблон оценок"


def test_xls_format_hint(tmp_path) -> None:
    old_file = tmp_path / "данные.xls"
    old_file.write_bytes(b"\xd0\xcf\x11\xe0")  # сигнатура OLE
    with pytest.raises(Exception) as exc:
        ExcelParser(str(old_file)).load()
    assert "сохраните его как .xlsx" in str(exc.value)


def _real_data_folder() -> str | None:
    """Каталог с настоящими файлами преподавателей (если он есть рядом с проектом)."""
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for candidate in (
        os.path.join(root, "real_data"),
        os.path.expanduser("~/Downloads/ДДЮТ оценки"),
    ):
        if os.path.isdir(candidate):
            return candidate
    return None


@pytest.mark.skipif(_real_data_folder() is None, reason="реальные файлы недоступны")
def test_real_files_are_parsed() -> None:
    """Дымовой тест на настоящих файлах преподавателей."""
    folder = _real_data_folder()
    assert folder is not None
    total_teachers = 0
    for name in sorted(os.listdir(folder)):
        if not name.endswith(".xlsx") or name.startswith("~") or "Админ" in name:
            continue
        parser = ExcelParser(os.path.join(folder, name), teacher_name=safe_file_stem(name))
        for teacher in parser.parse_all_teachers():
            assert teacher.directions, f"{name}: не найдено ни одного направления"
            for direction in teacher.directions:
                assert direction.evaluations, f"{name}: пустое направление {direction.name}"
                assert direction.total_score() > 0
            total_teachers += 1
    assert total_teachers >= 4
