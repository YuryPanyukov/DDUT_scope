"""Генерация иконки приложения: `app.ico` (для exe и окна) и `app.png`.

Иконка рисуется средствами самого PyQt5, поэтому дополнительных зависимостей
(например, Pillow) не требуется. Запуск:

    uv run python scripts/create_icon.py

Файлы попадают в `evaluation_app/resources/icons/` и коммитятся в репозиторий:
`.ico` нужен PyInstaller для exe, PNG — запасной вариант для платформ, где Qt
не умеет читать ICO.
"""

from __future__ import annotations

import math
import os
import struct
import sys
from pathlib import Path

from PyQt5.QtCore import QBuffer, QIODevice, QPointF, QRectF, Qt
from PyQt5.QtGui import (
    QBrush,
    QColor,
    QIcon,
    QImage,
    QLinearGradient,
    QPainter,
    QPolygonF,
)
from PyQt5.QtWidgets import QApplication

ROOT = Path(__file__).resolve().parent.parent
ICONS_DIR = ROOT / "evaluation_app" / "resources" / "icons"
ICO_FILE = ICONS_DIR / "app.ico"
PNG_FILE = ICONS_DIR / "app.png"

# Размеры внутри ICO: Windows берёт 16/32/48 для панели задач, 256 — для проводника.
ICO_SIZES = (16, 24, 32, 48, 64, 128, 256)
BASE_SIZE = 256  # геометрия описана в системе координат 256×256

BACKGROUND_TOP = QColor("#3b7ddd")
BACKGROUND_BOTTOM = QColor("#16386f")
PAPER = QColor("#ffffff")
LINE_COLOR = QColor("#5a8ed6")
STAR_FILL = QColor("#f5b301")
STAR_EDGE = QColor("#c98f00")
STAR_CENTER = (184.0, 182.0)
STAR_OUTER = 58.0
STAR_INNER = 24.0


def star_polygon(center_x: float, center_y: float, outer: float, inner: float) -> QPolygonF:
    """Пятиконечная звезда (луч вверх) — «оценка» в иконке."""
    points = []
    for index in range(10):
        radius = outer if index % 2 == 0 else inner
        angle = math.radians(-90.0 + index * 36.0)
        points.append(QPointF(center_x + radius * math.cos(angle), center_y + radius * math.sin(angle)))
    return QPolygonF(points)


def draw_icon(size: int) -> QImage:
    """Рисует иконку нужного размера: синий фон, лист оценок и звезда."""
    image = QImage(size, size, QImage.Format_ARGB32_Premultiplied)
    image.fill(Qt.transparent)

    painter = QPainter(image)
    painter.setRenderHint(QPainter.Antialiasing, True)
    scale = size / BASE_SIZE
    painter.scale(scale, scale)

    # Фон — скруглённый квадрат с градиентом.
    gradient = QLinearGradient(0.0, 0.0, float(BASE_SIZE), float(BASE_SIZE))
    gradient.setColorAt(0.0, BACKGROUND_TOP)
    gradient.setColorAt(1.0, BACKGROUND_BOTTOM)
    painter.setPen(Qt.NoPen)
    painter.setBrush(QBrush(gradient))
    painter.drawRoundedRect(QRectF(8.0, 8.0, 240.0, 240.0), 56.0, 56.0)

    # Лист с оценками.
    painter.setBrush(PAPER)
    painter.drawRoundedRect(QRectF(64.0, 44.0, 128.0, 168.0), 16.0, 16.0)

    # Строки списка.
    painter.setBrush(LINE_COLOR)
    for index, width in enumerate((84.0, 84.0, 56.0)):
        y = 88.0 + index * 36.0
        painter.drawRoundedRect(QRectF(88.0, y, width, 16.0), 8.0, 8.0)

    # Звезда поверх правого нижнего угла листа.
    painter.setPen(STAR_EDGE)
    painter.setBrush(STAR_FILL)
    painter.drawPolygon(star_polygon(*STAR_CENTER, STAR_OUTER, STAR_INNER))

    painter.end()
    return image


def png_bytes(image: QImage) -> bytes:
    """PNG-представление картинки."""
    buffer = QBuffer()
    buffer.open(QIODevice.WriteOnly)
    image.save(buffer, "PNG")
    return bytes(buffer.data())


def write_ico(path: Path, images: dict[int, bytes]) -> None:
    """Собирает ICO-файл: заголовок, таблица записей и сами PNG-картинки.

    Формат ICO допускает PNG внутри записи (Windows Vista и новее), поэтому
    обходимся без сторонних библиотек.
    """
    header = struct.pack("<HHH", 0, 1, len(images))
    offset = len(header) + 16 * len(images)

    entries, blobs = [], []
    for size, data in images.items():
        # Для 256 в поле размера пишется 0 — таково соглашение формата.
        dimension = 0 if size >= 256 else size
        entries.append(struct.pack("<BBBBHHII", dimension, dimension, 0, 0, 1, 32, len(data), offset))
        blobs.append(data)
        offset += len(data)

    path.write_bytes(header + b"".join(entries) + b"".join(blobs))


def main() -> int:
    # До создания QApplication: иконка рисуется без экрана.
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    # Ссылку на приложение обязательно сохраняем: без неё Qt-объекты (QIcon,
    # QPixmap) отказываются работать — "Must construct a QGuiApplication".
    application = QApplication(sys.argv)
    assert application is not None

    ICONS_DIR.mkdir(parents=True, exist_ok=True)

    blobs = {size: png_bytes(draw_icon(size)) for size in ICO_SIZES}
    write_ico(ICO_FILE, blobs)
    PNG_FILE.write_bytes(blobs[BASE_SIZE])

    print(f"Иконка: {ICO_FILE} ({ICO_FILE.stat().st_size} байт, размеры {list(ICO_SIZES)})")
    print(f"PNG:    {PNG_FILE} ({PNG_FILE.stat().st_size} байт)")

    icon = QIcon(str(ICO_FILE))
    print("Загружена обратно, размеры:", [f"{s.width()}x{s.height()}" for s in icon.availableSizes()])
    return 0 if not icon.isNull() else 1


if __name__ == "__main__":
    raise SystemExit(main())
