# План разработки приложения: Сбор и обработка оценок преподавателей

---

## 1. Общее описание проекта

| Параметр | Значение |
|---|---|
| **Название** | Оценочный комитет (ОценкиПреподавателей) |
| **Язык** | Python 3.10+ |
| **GUI-фреймворк** | PyQt5 (или PyQt6) |
| **Библиотека для Excel** | openpyxl |
| **Хранение данных** | JSON-файл (сохранение проекта) |
| **Тип приложения** | Desktop-приложение для администратора |

---

## 2. Архитектура приложения

```
┌─────────────────────────────────────────────────────┐
│                   PyQt5 GUI                         │
│  ┌──────────┐ ┌──────────────┐ ┌────────────────┐   │
│  │Главная   │ │ Загрузка     │ │ Таблица        │   │
│  │форма     │ │ Excel-файлы  │ │ результатов    │   │
│  └──────────┘ └──────────────┘ └────────────────┘   │
│  ┌──────────┐ ┌──────────────┐ ┌────────────────┐   │
│  │Список    │ │ Сортировка   │ │ Сохранение/    │   │
│  │препода-  │ │ по ФИО/баллам│ │ Открытие       │   │
│  │телей     │ │              │ │ проекта        │   │
│  └──────────┘ └──────────────┘ └────────────────┘   │
├─────────────────────────────────────────────────────┤
│              Бизнес-логика (Service Layer)          │
│  ┌──────────────┐ ┌─────────────┐ ┌──────────────┐ │
│  │ ExcelParser  │ │ DataMerger  │ │ DataSorter   │ │
│  │ (openpyxl)   │ │ (суммация)  │ │ (сортировка) │ │
│  └──────────────┘ └─────────────┘ └──────────────┘ │
├─────────────────────────────────────────────────────┤
│               Модель данных (Model)                 │
│  ┌──────────────┐ ┌─────────────┐ ┌──────────────┐ │
│  │ Teacher      │ │ Student     │ │ Direction    │ │
│  │ (преподаватель)│ │ (ученик)  │ │ (направление)│ │
│  └──────────────┘ └─────────────┘ └──────────────┘ │
├─────────────────────────────────────────────────────┤
│            Проект (Project State)                   │
│  Список преподавателей → их направления → оценки   │
└─────────────────────────────────────────────────────┘
```

---

## 3. Модель данных

### 3.1. Структура данных

```python
# Основная структура проекта
Project = {
    "project_name": str,
    "academic_year": str,
    "teachers": [
        Teacher
    ]
}

Teacher = {
    "id": UUID/str,
    "name": str,              # ФИО преподавателя
    "directions": [
        Direction
    ]
}

Direction = {
    "id": UUID/str,
    "name": str,              # Название направления (лист Excel)
    "evaluations": [
        Evaluation
    ]
}

Evaluation = {
    "id": UUID/str,
    "student_name": str,      # ФИО ученика
    "competition_name": str,   # Название конкурса
    "total_score": float,      # Итого (сумма баллов)
}
```

### 3.2. Агрегированная таблица (для отображения администратору)

```python
AggregatedRecord = {
    "student_name": str,          # ФИО ученика
    "total_score": float,          # Сумма по всем преподавателям и направлениям
    "teachers_involved": [str],    # Какие преподаватели оценивали
    "directions": [str]            # В каких направлениях
}
```

---

## 4. Структура файлов проекта

```
evaluation_app/
├── main.py                      # Точка входа
├── config.py                    # Конфигурация (пути, настройки)
├── requirements.txt             # Зависимости
│
├── models/
│   ├── __init__.py
│   ├── project.py               # Класс Project
│   ├── teacher.py               # Класс Teacher
│   ├── direction.py             # Класс Direction
│   └── evaluation.py            # Класс Evaluation
│
├── services/
│   ├── __init__.py
│   ├── excel_parser.py          # Парсинг Excel (openpyxl)
│   ├── data_merger.py           # Мерджинг данных (суммация ФИО)
│   ├── data_sorter.py           # Сортировка данных
│   └── project_manager.py       # Сохранение/загрузка проекта (JSON)
│
├── ui/
│   ├── __init__.py
│   ├── main_window.py           # Главное окно
│   ├── upload_dialog.py         # Диалог загрузки файлов
│   ├── teacher_list_widget.py   # Список преподавателей
│   ├── direction_view.py        # Просмотр направления
│   ├── summary_table_view.py    # Общая таблица результатов
│   ├── sort_widget.py           # Виджет сортировки
│   └── settings_dialog.py       # Настройки (путь сохранения)
│
├── utils/
│   ├── __init__.py
│   ├── file_utils.py            # Работа с файлами
│   └── validators.py            # Валидация данных
│
└── resources/
    └── icons/                   # Иконки (при необходимости)
```

---

## 5. Подробный план разработки по этапам

### Этап 1: Подготовка (1-2 дня)

| Задача | Описание |
|---|---|
| Определить формат Excel-файлов | Создать шаблоны таблиц для преподавателей (заголовки, структура листов) |
| Подготовить тестовые данные | Создать 2-3 тестовых Excel-файла с разными сценариями |
| Настроить окружение | `venv`, `requirements.txt`, базовая структура проекта |
| Создать шаблон Excel | Таблица с колонками: ФИО, Название конкурса, Итого |

**requirements.txt:**
```
PyQt5>=5.15.0
openpyxl>=3.0.0
```

**Шаблон Excel (пример структуры одного листа):**

| ФИО | Название конкурса | Итого |
|---|---|---|
| Иванов И.И. | Олимпиада по математике | 85 |
| Петрова А.С. | Олимпиада по математике | 92 |
| ... | ... | ... |

---

### Этап 2: Модель данных (2-3 дня)

**Задача:** Реализовать классы модели данных и сериализацию.

```python
# models/evaluation.py
from dataclasses import dataclass, field
from uuid import uuid4

@dataclass
class Evaluation:
    id: str = field(default_factory=lambda: str(uuid4()))
    student_name: str = ""
    competition_name: str = ""
    total_score: float = 0.0

# models/direction.py
@dataclass
class Direction:
    id: str = field(default_factory=lambda: str(uuid4()))
    name: str = ""
    evaluations: list[Evaluation] = field(default_factory=list)

# models/teacher.py
@dataclass
class Teacher:
    id: str = field(default_factory=lambda: str(uuid4()))
    name: str = ""
    directions: list[Direction] = field(default_factory=list)

# models/project.py
@dataclass
class Project:
    project_name: str = ""
    academic_year: str = ""
    teachers: list[Teacher] = field(default_factory=list)
```

**Тестирование:** Юнит-тесты для каждого класса — создание объектов, сериализация, десериализация.

---

### Этап 3: Парсинг Excel (3-4 дня)

**Задача:** Реализовать `ExcelParser` с использованием `openpyxl`.

```python
# services/excel_parser.py
import openpyxl
from models.teacher import Teacher
from models.direction import Direction
from models.evaluation import Evaluation
from typing import list

class ExcelParser:
    """Парсер Excel-файлов преподавателей"""

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.workbook = None

    def load(self):
        """Открытие Excel-файла"""
        self.workbook = openpyxl.load_workbook(self.file_path)

    def parse_all_teachers(self) -> list[Teacher]:
        """
        Парсинг всех листов как направлений.
        Считаем, что каждый лист — направление одного преподавателя,
        либо ФИО преподавателя определяется из метаданных.
        """
        teachers = []
        for sheet_name in self.workbook.sheetnames:
            worksheet = self.workbook[sheet_name]
            direction = self._parse_sheet(sheet_name, worksheet)
            # Здесь нужно определить логику привязки к преподавателю
            teachers.append(...)
        return teachers

    def _parse_sheet(self, sheet_name: str, worksheet) -> Direction:
        """
        Парсинг одного листа Excel.
        Ожидаемая структура:
          - Столбец A: ФИО
          - Столбец B: Название конкурса
          - Столбец C: Итого
        """
        direction = Direction(name=sheet_name)

        # Определяем заголовки (подстройка под реальный формат)
        for row in range(1, worksheet.max_row + 1):
            fio = worksheet.cell(row=row, column=1).value
            competition = worksheet.cell(row=row, column=2).value
            total = worksheet.cell(row=row, column=3).value

            if fio and competition:
                evaluation = Evaluation(
                    student_name=str(fio).strip(),
                    competition_name=str(competition).strip(),
                    total_score=float(total) if total else 0.0
                )
                direction.evaluations.append(evaluation)

        return direction

    def get_sheet_names(self) -> list[str]:
        """Получение списка листов"""
        if not self.workbook:
            self.load()
        return self.workbook.sheetnames
```

**Обработка неоднозначностей:**
- Если один файл содержит несколько преподавателей — добавить диалоговое окно для маппинга листов к преподавателям
- Вариант: один файл = один преподаватель, каждый лист = его направление (наиболее вероятный сценарий)

**Тестирование:** Парсинг тестовых файлов, проверка корректности чтения разных структур.

---

### Этап 4: Бизнес-логика обработки (3-4 дня)

#### 4.1. DataMerger — Агрегация данных

```python
# services/data_merger.py
from collections import defaultdict
from models.project import Project
from models.evaluation import Evaluation

class DataMerger:
    """Объединение данных учеников с одинаковым ФИО"""

    @staticmethod
    def merge_by_student(project: Project) -> list[dict]:
        """
        Формирует агрегированную таблицу:
        ключ — ФИО ученика, значение — сумма баллов по всем
        преподавателям и направлениям.
        """
        student_scores = defaultdict(lambda: {
            "total_score": 0.0,
            "teachers": [],
            "directions": [],
            "competitions": []
        })

        for teacher in project.teachers:
            for direction in teacher.directions:
                for eval_item in direction.evaluations:
                    key = eval_item.student_name
                    student_scores[key]["total_score"] += eval_item.total_score
                    student_scores[key]["teachers"].append(teacher.name)
                    student_scores[key]["directions"].append(direction.name)
                    student_scores[key]["competitions"].append(
                        eval_item.competition_name
                    )

        result = []
        for fio, data in student_scores.items():
            result.append({
                "student_name": fio,
                "total_score": data["total_score"],
                "teachers_involved": list(set(data["teachers"])),
                "directions": list(set(data["directions"])),
                "competitions": data["competitions"]
            })

        return result
```

#### 4.2. DataSorter — Сортировка

```python
# services/data_sorter.py
class DataSorter:
    """Сортировка агрегированных данных"""

    @staticmethod
    def sort_by_name(data: list[dict], reverse: bool = False) -> list[dict]:
        """Сортировка по ФИО (по алфавиту)"""
        return sorted(data, key=lambda x: x["student_name"], reverse=reverse)

    @staticmethod
    def sort_by_score(data: list[dict], reverse: bool = True) -> list[dict]:
        """Сортировка по баллам (по убыванию по умолчанию)"""
        return sorted(data, key=lambda x: x["total_score"], reverse=reverse)
```

#### 4.3. ProjectManager — Сохранение/Загрузка

```python
# services/project_manager.py
import json
import os
from models.project import Project

class ProjectManager:
    """Управление сохранением и загрузкой проекта"""

    @staticmethod
    def save_project(project: Project, file_path: str):
        """Сохранение проекта в JSON"""
        data = {
            "project_name": project.project_name,
            "academic_year": project.academic_year,
            "teachers": [
                {
                    "id": t.id,
                    "name": t.name,
                    "directions": [
                        {
                            "id": d.id,
                            "name": d.name,
                            "evaluations": [
                                {
                                    "id": e.id,
                                    "student_name": e.student_name,
                                    "competition_name": e.competition_name,
                                    "total_score": e.total_score,
                                }
                                for e in d.evaluations
                            ]
                        }
                        for d in t.directions
                    ]
                }
                for t in project.teachers
            ]
        }
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    @staticmethod
    def load_project(file_path: str) -> Project:
        """Загрузка проекта из JSON"""
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        project = Project(
            project_name=data["project_name"],
            academic_year=data.get("academic_year", "")
        )

        for t_data in data.get("teachers", []):
            teacher = Teacher(id=t_data["id"], name=t_data["name"])
            for d_data in t_data.get("directions", []):
                direction = Direction(id=d_data["id"], name=d_data["name"])
                for e_data in d_data.get("evaluations", []):
                    evaluation = Evaluation(
                        id=e_data["id"],
                        student_name=e_data["student_name"],
                        competition_name=e_data["competition_name"],
                        total_score=e_data["total_score"],
                    )
                    direction.evaluations.append(evaluation)
                teacher.directions.append(direction)
            project.teachers.append(teacher)

        return project
```

**Обновление данных преподавателя:**

```python
# В ProjectManager или сервисе
@staticmethod
def update_teacher(project: Project, teacher_id: str, new_teacher_data: Teacher):
    """Обновление данных конкретного преподавателя (удаление старых + вставка новых)"""
    for i, teacher in enumerate(project.teachers):
        if teacher.id == teacher_id:
            project.teachers[i] = new_teacher_data
            return True
    return False
```

---

### Этап 5: GUI-уровень — PyQt5 (10-14 дней)

#### 5.1. Главное окно (`main_window.py`)

```
┌──────────────────────────────────────────────────────────────┐
│ Оценочный комитет                    [Сохранить] [Загрузить] │
├──────────────┬───────────────────────────────────────────────┤
│              │                                               │
│  Список      │  [ФИО ▼]  [Баллы ▼]    [Обновить]           │
│  преподават- │                                               │
│ елей         │  ┌─────────────────────────────────────┐     │
│  ▸ Петрова   │  │ Сводная таблица:                    │     │
│  ▸ Иванов    │  │ ФИО        │ Итого  │ Конкурсы     │     │
│  ▸ Сидорова  │  │ Иванов И.  │ 177    │ Math, Sci    │     │
│  ▸ ...       │  │ Петрова А. │ 92     │ Math         │     │
│              │  │ ...        │ ...    │ ...          │     │
│  [Добавить]  │  └─────────────────────────────────────┘     │
│  [Удалить]   │                                               │
│              │  Детальный просмотр (при выборе преподавателя)│
│              │  ┌─────────────────────────────────────┐     │
│              │  │ Преподаватель: Петрова А.           │     │
│              │  │ Направление: Естественные науки      │     │
│              │  │ ФИО      │ Конкурс  │ Баллы         │     │
│              │  └─────────────────────────────────────┘     │
└──────────────┴───────────────────────────────────────────────┘
```

#### 5.2. Ключевые UI-компоненты

**Главное окно — основные виджеты:**

```python
# ui/main_window.py
from PyQt5.QtWidgets import (
    QMainWindow, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QHBoxLayout, QWidget, QPushButton,
    QFileDialog, QHeaderView, QComboBox, QLabel,
    QLineEdit, QMessageBox, QGroupBox
)
from PyQt5.QtCore import Qt
from services.excel_parser import ExcelParser
from services.data_merger import DataMerger
from services.data_sorter import DataSorter
from services.project_manager import ProjectManager
from models.project import Project

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Оценочный комитет")
        self.setGeometry(100, 100, 1100, 700)
        self.project = Project()
        self.current_project_path = None
        self._init_ui()

    def _init_ui(self):
        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # === Левая панель: Список преподавателей ===
        left_group = QGroupBox("Преподаватели")
        left_layout = QVBoxLayout(left_group)
        self.teacher_list_widget = QListWidget()  # см. ниже
        left_layout.addWidget(self.teacher_list_widget)
        btn_add_teacher = QPushButton("Добавить")
        btn_remove_teacher = QPushButton("Удалить")
        left_layout.addWidget(btn_add_teacher)
        left_layout.addWidget(btn_remove_teacher)

        # === Правая панель ===
        right_layout = QVBoxLayout()

        # Верхняя панель: кнопки управления
        top_bar = QHBoxLayout()
        btn_load_excel = QPushButton("Загрузить Excel")
        btn_save = QPushButton("Сохранить проект")
        btn_open = QPushButton("Открыть проект")
        btn_update = QPushButton("Обновить данные")
        btn_sort_name = QPushButton("Сортировать по ФИО")
        btn_sort_score = QPushButton("Сортировать по баллам")
        top_bar.addWidget(btn_load_excel)
        top_bar.addWidget(btn_save)
        top_bar.addWidget(btn_open)
        top_bar.addWidget(btn_update)
        top_bar.addWidget(btn_sort_name)
        top_bar.addWidget(btn_sort_score)
        right_layout.addLayout(top_bar)

        # Сводная таблица
        self.summary_table = QTableWidget()
        self.summary_table.setAlternatingRowColors(True)
        right_layout.addWidget(self.summary_table)

        # Детальный просмотр (для выбранного преподавателя)
        self.detail_table = QTableWidget()
        right_layout.addWidget(self.detail_table)

        main_layout.addWidget(left_group)
        main_layout.addLayout(right_layout)

        # Связь сигналов со слотами
        btn_load_excel.clicked.connect(self._on_load_excel)
        btn_save.clicked.connect(self._on_save_project)
        btn_open.clicked.connect(self._on_open_project)
        btn_update.clicked.connect(self._on_update_teacher)
        btn_sort_name.clicked.connect(self._on_sort_by_name)
        btn_sort_score.clicked.connect(self._on_sort_by_score)
        self.teacher_list_widget.currentRowChanged.connect(
            self._on_select_teacher
        )

    # === Обработчики событий ===

    def _on_load_excel(self):
        """Загрузка одного или нескольких Excel-файлов"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Выберите Excel-файлы", "",
            "Excel files (*.xlsx);;All files (*)"
        )
        if not file_paths:
            return

        for path in file_paths:
            parser = ExcelParser(path)
            parser.load()
            # Парсинг и добавление преподавателей
            teachers = parser.parse_all_teachers()
            for teacher in teachers:
                self.project.teachers.append(teacher)

        self._refresh_teacher_list()
        self._refresh_summary_table()

    def _on_save_project(self):
        """Сохранение проекта"""
        if not self.current_project_path:
            path, _ = QFileDialog.getSaveFileName(
                self, "Сохранить проект", "",
                "JSON files (*.json);;All files (*)"
            )
            if not path:
                return
            self.current_project_path = path

        ProjectManager.save_project(self.project, self.current_project_path)
        QMessageBox.information(self, "Сохранение", "Проект сохранён!")

    def _on_open_project(self):
        """Загрузка ранее сохранённого проекта"""
        path, _ = QFileDialog.getOpenFileName(
            self, "Открыть проект", "",
            "JSON files (*.json);;All files (*)"
        )
        if not path:
            return
        self.current_project_path = path
        self.project = ProjectManager.load_project(path)
        self._refresh_teacher_list()
        self._refresh_summary_table()

    def _on_update_teacher(self):
        """Обновление данных конкретного преподавателя"""
        row = self.teacher_list_widget.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите преподавателя!")
            return

        teacher = self.project.teachers[row]
        file_path, _ = QFileDialog.getOpenFileName(
            self, f"Выберите файл для {teacher.name}", "",
            "Excel files (*.xlsx);;All files (*)"
        )
        if not file_path:
            return

        parser = ExcelParser(file_path)
        parser.load()
        new_teachers = parser.parse_all_teachers()
        if new_teachers:
            # Замена: удаляем старые данные, вставляем новые
            self.project.teachers[row] = new_teachers[0]
            self._refresh_teacher_list()
            self._refresh_summary_table()

    def _on_sort_by_name(self):
        self._refresh_summary_table(sort_by="name")

    def _on_sort_by_score(self):
        self._refresh_summary_table(sort_by="score")

    def _on_select_teacher(self, row):
        """Просмотр данных конкретного преподавателя"""
        if row >= 0:
            self._refresh_detail_table(self.project.teachers[row])

    # === Обновление представлений ===

    def _refresh_teacher_list(self):
        self.teacher_list_widget.clear()
        for teacher in self.project.teachers:
            self.teacher_list_widget.addItem(teacher.name)

    def _refresh_summary_table(self, sort_by=None):
        data = DataMerger.merge_by_student(self.project)

        if sort_by == "name":
            data = DataSorter.sort_by_name(data)
        elif sort_by == "score":
            data = DataSorter.sort_by_score(data)

        self.summary_table.setRowCount(len(data))
        self.summary_table.setColumnCount(3)
        self.summary_table.setHorizontalHeaderLabels(["ФИО", "Итого баллов", "Конкурсы"])

        for i, record in enumerate(data):
            self.summary_table.setItem(i, 0, QTableWidgetItem(record["student_name"]))
            self.summary_table.setItem(i, 1, QTableWidgetItem(str(record["total_score"])))
            competitions = ", ".join(record.get("competitions", []))
            self.summary_table.setItem(i, 2, QTableWidgetItem(competitions))

        self.summary_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)

    def _refresh_detail_table(self, teacher):
        self.detail_table.clear()
        row_offset = 0
        for direction in teacher.directions:
            self.detail_table.setRowCount(len(direction.evaluations))
            self.detail_table.setColumnCount(3)
            self.detail_table.setHorizontalHeaderLabels(["ФИО", "Конкурс", "Баллы"])
            for j, eval_item in enumerate(direction.evaluations):
                self.detail_table.setItem(row_offset + j, 0,
                    QTableWidgetItem(eval_item.student_name))
                self.detail_table.setItem(row_offset + j, 1,
                    QTableWidgetItem(eval_item.competition_name))
                self.detail_table.setItem(row_offset + j, 2,
                    QTableWidgetItem(str(eval_item.total_score)))
            row_offset += len(direction.evaluations)
```

---

### Этап 6: Диалог загрузки и настройки (3-4 дня)

#### 6.1. Диалог привязки листов к преподавателям

Когда загружается один Excel-файл с несколькими листами, необходимо определить, кто преподаватель. Реализовать диалог:

```
Загружен файл: "Петрова_2024.xlsx"
┌─────────────────────────────────────────┐
│ Лист              │ Преподаватель       │
├───────────────────┼─────────────────────┤
│ Естественные науки│ [Петрова А.И. ▼]    │
│ Математика        │ [Петрова А.И. ▼]    │
│ Информатика       │ [Петрова А.И. ▼]    │
└───────────────────┴─────────────────────┘
[Отмена] [Готово]
```

#### 6.2. Диалог создания/настройки проекта

```python
# ui/settings_dialog.py
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit

class ProjectSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Новый проект")
        layout = QVBoxLayout(self)

        h_layout1 = QHBoxLayout()
        h_layout1.addWidget(QLabel("Название проекта:"))
        self.project_name_edit = QLineEdit()
        h_layout1.addWidget(self.project_name_edit)
        layout.addLayout(h_layout1)

        h_layout2 = QHBoxLayout()
        h_layout2.addWidget(QLabel("Учебный год:"))
        self.year_edit = QLineEdit("2024-2025")
        h_layout2.addWidget(self.year_edit)
        layout.addLayout(h_layout2)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        ok_btn = QPushButton("ОК")
        ok_btn.clicked.connect(self.accept)
        btn_layout.addWidget(ok_btn)
        layout.addLayout(btn_layout)

    def get_project_name(self):
        return self.project_name_edit.text()

    def get_academic_year(self):
        return self.year_edit.text()
```

#### 6.3. Настройка пути сохранения по умолчанию

```python
# config.py
import os

DEFAULT_SAVE_PATH = os.path.join(os.path.expanduser("~"), "evaluation_projects")
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".eval_app_settings.json")
```

---

### Этап 7: Обработка ошибок и валидация (2-3 дня)

| Сценарий | Обработка |
|---|---|
| Файл Excel не найден | `QMessageBox.critical()` с описанием ошибки |
| Неверный формат таблицы (нет нужных колонок) | Предупреждение + предложить переназначить столбцы |
| Пустые ячейки | Пропустить строку или записать значение по умолчанию |
| Некорректные данные (текст вместо числа в "Итого") | `try/except` + диалог с пользователем |
| Дублирование ФИО внутри одного листа | Автоматическая суммация или предупреждение |
| Попытка открыть несуществующий JSON-проект | Обработка `FileNotFoundError` |
| Нечётное количество столбцов в Excel | Диалоговое окно для ручного указания структуры |

```python
# utils/validators.py
class DataValidator:
    @staticmethod
    def validate_excel_headers(worksheet) -> bool:
        """Проверка наличия ожидаемых заголовков"""
        expected = {"ФИО", "Название конкурса", "Итого"}
        headers = set()
        for col in range(1, worksheet.max_column + 1):
            val = worksheet.cell(row=1, column=col).value
            if val:
                headers.add(str(val).strip())
        return expected.issubset(headers)

    @staticmethod
    def validate_score(value):
        """Проверка, что значение баллов — число"""
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False
```

---

### Этап 8: Полировка и тестирование (3-5 дней)

#### 8.1. Юнит-тесты

```python
# tests/test_excel_parser.py
import unittest
from unittest.mock import MagicMock
from services.excel_parser import ExcelParser

class TestExcelParser(unittest.TestCase):
    def test_parse_single_sheet(self):
        """Парсинг файла с одним листом"""
        # ... mock-объект worksheet с тестовыми данными ...

    def test_parse_multi_sheet(self):
        """Парсинг файла с несколькими листами"""

    def test_empty_rows_are_skipped(self):
        """Пустые строки пропускаются"""

    def test_invalid_scores_handled(self):
        """Некорректные баллы обрабатываются"""

# tests/test_data_merger.py
class TestDataMerger(unittest.TestCase):
    def test_merge_duplicate_fio(self):
        """Суммация одинаковых ФИО из разных преподавателей"""

    def test_merge_no_duplicates(self):
        """Все ФИО уникальны"""

    def test_merge_empty_project(self):
        """Пустой проект"""

# tests/test_project_manager.py
class TestProjectManager(unittest.TestCase):
    def test_save_and_load_cycle(self):
        """Сохранение и загрузка"""

    def test_update_teacher(self):
        """Обновление данных преподавателя"""
```

#### 8.2. Интеграционное тестирование

1. Загрузить 3 Excel-файла → проверить, что преподаватели созданы
2. Нажать «Сохранить» → закрыть → «Открыть» → данные идентичны
3. Выбрать преподавателя → «Обновить» → старые данные удалены, новые отображаются
4. Сортировка по ФИО — проверка алфавитного порядка
5. Сортировка по баллам — проверка убывания
6. Дублирующиеся ФИО → суммация корректна

---

## 6. Сводная таблица ключевых классов и функций

| Класс/Модуль | Назначение | Ключевые методы |
|---|---|---|
| `Project` (model) | Контейнер данных проекта | `add_teacher()`, `remove_teacher()` |
| `Teacher` (model) | Данные преподавателя | `add_direction()`, `remove_direction()` |
| `Direction` (model) | Направление (лист Excel) | `add_evaluation()` |
| `Evaluation` (model) | Оценка ученика | — |
| `ExcelParser` | Чтение .xlsx | `load()`, `parse_all_teachers()`, `_parse_sheet()` |
| `DataMerger` | Агрегация данных | `merge_by_student()` |
| `DataSorter` | Сортировка | `sort_by_name()`, `sort_by_score()` |
| `ProjectManager` | Сохранение/загрузка | `save_project()`, `load_project()`, `update_teacher()` |
| `MainWindow` (UI) | Главный интерфейс | `_on_load_excel()`, `_on_save_project()`, `_on_update_teacher()` |
| `DataValidator` | Валидация | `validate_excel_headers()`, `validate_score()` |

---

## 7. Порядок реализации (дорожная карта)

```
Неделя 1:
  ✅ Этап 1: Подготовка (шаблон Excel, тестовые данные, venv)
  ✅ Этап 2: Модель данных (dataclass-ы, тесты)

Неделя 2:
  ✅ Этап 3: Парсинг Excel (openpyxl)
  ✅ Начало Этапа 4: DataMerger, DataSorter

Неделя 3:
  ✅ Этап 4: ProjectManager (сохранение/загрузка)
  ✅ Этап 5: GUI — Главное окно, загрузка файлов

Неделя 4:
  ✅ Этап 5: GUI — Таблицы, сортировка, просмотр деталей
  ✅ Этап 6: Диалоги (настройки, привязка листов)

Неделя 5:
  ✅ Этап 7: Обработка ошибок и валидация
  ✅ Этап 8: Тестирование, полировка
```

---

## 8. Дополнительные рекомендации

1. **Поддержка шаблонов Excel разных форматов** — реализовать гибкое определение столбцов (заголовки могут называться по-разному: «ФИО» / «Фамилия Имя Отчество» / «Ученик»).

2. **Предпросмотр перед загрузкой** — показать пользователю содержимое Excel перед фактическим добавлением в проект.

3. **Экспорт итоговой таблицы в Excel** — удобная функция для администратора: `openpyxl` может не только читать, но и записывать.

4. **История изменений** — при обновлении данных преподавателя можно хранить предыдущую версию (откат).

5. **Поиск по ФИО** — добавить строку поиска для быстрой фильтрации в сводной таблице.

6. **Статистика** — минимальные/максимальные баллы, количество учеников, количество конкурсов.

---

Этот план покрывает все 7 пунктов из вашего ТЗ и даёт чёткую дорожную карту для пошаговой реализации приложения.