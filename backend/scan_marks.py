"""Отметки в квадратах (☒ / ☐) на сканах.

Tesseract не различает пустой и отмеченный квадрат: «☒ Второй инспекционный контроль»
превращается в «С Второй…», «☐ Первый…» — в «[]», «С]» или пропадает. Поэтому квадраты
ищутся на картинке страницы: квадрат — это горизонтальные отрезки сверху и снизу одной длины
и вертикальные по краям; отмечен он, если внутри есть чернила (крест, галочка).

Только Pillow, без numpy: строки картинки разбираются регулярными выражениями по байтам.
"""
import re
from typing import Callable, Optional

from PIL import Image

_DARK = 160          # пиксель темнее — чернила (синяя ручка на скане серее чёрного шрифта)
_FILL_CHECKED = 0.06  # доля чернил внутри квадрата, с которой он считается отмеченным


def _binary_rows(img: Image.Image) -> list:
    """Строки картинки байтами: 0 — чернила, 255 — фон."""
    bw = img.convert("L").point(lambda v: 0 if v < _DARK else 255)
    w, h = bw.size
    data = bw.tobytes()
    return [data[y * w:(y + 1) * w] for y in range(h)]


def _runs(row: bytes, lo: int, hi: int) -> list:
    """Отрезки чернил длиной lo..hi (концы — с допуском в 1 пиксель фона внутри отрезка)."""
    rx = re.compile(rb"(?<!\x00)\x00(?:\x00|\xff(?=\x00)){%d,%d}(?!\x00)" % (lo - 1, hi - 1))
    return [(m.start(), m.end()) for m in rx.finditer(row)]


def _dark_share(rows: list, x0: int, y0: int, x1: int, y1: int) -> float:
    if x1 <= x0 or y1 <= y0:
        return 0.0
    dark = sum(rows[y][x0:x1].count(0) for y in range(y0, y1))
    return dark / ((x1 - x0) * (y1 - y0))


def _column_dark(rows: list, x: int, y0: int, y1: int, width: int = 3) -> float:
    """Доля строк y0..y1, где в колонках x..x+width есть чернила (вертикальная сторона)."""
    if y1 <= y0:
        return 0.0
    hit = sum(1 for y in range(y0, y1) if 0 in rows[y][max(0, x):x + width])
    return hit / (y1 - y0)


def find_boxes(img: Image.Image, dpi: int) -> list:
    """Квадраты-отметки на странице: [{"box": (x0, y0, x1, y1), "fill": доля чернил внутри,
    "checked": bool}] в порядке чтения (сверху вниз, слева направо)."""
    rows = _binary_rows(img)
    lo, hi = int(dpi * 0.1), int(dpi * 0.22)        # сторона квадрата ≈ 2,5–5,5 мм
    tops = {y: _runs(r, lo, hi) for y, r in enumerate(rows)}
    found = []
    height = len(rows)
    for y, runs in tops.items():
        for x0, x1 in runs:
            w = x1 - x0
            for h in range(int(w * 0.8), int(w * 1.25) + 1):
                yb = y + h
                if yb >= height:
                    break
                if not any(abs(a - x0) <= 3 and abs(b - x1) <= 3 for a, b in tops.get(yb, ())):
                    continue
                if _column_dark(rows, x0, y, yb) < 0.85 or _column_dark(rows, x1 - 3, y, yb) < 0.85:
                    continue
                found.append((x0, y, x1, yb))
                break
    # Толщина линий даёт по несколько находок на квадрат — оставляем внешнюю рамку
    boxes = []
    for b in sorted(found, key=lambda b: (b[2] - b[0]) * (b[3] - b[1]), reverse=True):
        if any(_overlap(b, k) > 0.4 for k in boxes):
            continue
        boxes.append(b)
    result = []
    for x0, y0, x1, y1 in boxes:
        ring = max(4, (x1 - x0) // 4)
        # Квадрат стоит отдельно: слева от него пусто. Жирная буква заголовка («Ч», «В»),
        # похожая на квадрат, касается соседних букв. Справа и сверху не проверяем —
        # рукописная галочка выходит за рамку вверх-вправо.
        if _dark_share(rows, max(0, x0 - ring), y0 + 2, x0 - 1, y1 - 2) > 0.02:
            continue
        pad = max(4, (x1 - x0) // 6)                 # отступ от рамки внутрь
        fill = _dark_share(rows, x0 + pad, y0 + pad, x1 - pad, y1 - pad)
        if fill > 0.7:
            continue  # сплошной чёрный прямоугольник — не квадрат для отметки
        result.append({"box": (x0, y0, x1, y1), "fill": round(fill, 3), "checked": fill >= _FILL_CHECKED})
    return _reading_order(result)


def _overlap(a: tuple, b: tuple) -> float:
    ix = max(0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0, min(a[3], b[3]) - max(a[1], b[1]))
    inter = ix * iy
    smaller = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
    return inter / smaller if smaller else 0.0


def _reading_order(items: list) -> list:
    """Сортировка по строкам: квадраты, чьи центры по высоте ближе половины стороны, — одна строка."""
    items = sorted(items, key=lambda i: (i["box"][1] + i["box"][3]) / 2)
    lines, current = [], []
    for it in items:
        cy = (it["box"][1] + it["box"][3]) / 2
        side = it["box"][3] - it["box"][1]
        if current and abs(cy - current[-1]["cy"]) > side * 0.6:
            lines.append(current)
            current = []
        current.append({**it, "cy": cy})
    if current:
        lines.append(current)
    out = []
    for n, line in enumerate(lines):
        for it in sorted(line, key=lambda i: i["box"][0]):
            it["line"] = n
            it["in_line"] = len(line)
            out.append(it)
    return out


def label_marks(img: Image.Image, boxes: list, ocr: Callable[[Image.Image], str]) -> list:
    """[(отмечено, подпись)] для секции «=== ОТМЕТКИ ===».

    Подпись — текст справа от квадрата до следующего квадрата в строке (ocr — распознавание
    вырезанного куска). Строка из четырёх и больше квадратов — таблица оценок (Лист оценки):
    у её клеток подписей нет, им даётся общая подпись «строка таблицы оценок N»."""
    width = img.size[0]
    marks, grid_rows = [], 0
    by_line = {}
    for b in boxes:
        by_line.setdefault(b["line"], []).append(b)
    for line in sorted(by_line):
        items = by_line[line]
        if len(items) >= 4:
            grid_rows += 1
            marks += [(b["checked"], f"строка таблицы оценок {grid_rows}") for b in items]
            continue
        for k, b in enumerate(items):
            x0, y0, x1, y1 = b["box"]
            side = y1 - y0
            right = items[k + 1]["box"][0] - side // 2 if k + 1 < len(items) else width - side
            crop = img.crop((x1 + side // 3, max(0, y0 - side), right, min(img.size[1], y1 + side)))
            label = re.sub(r"\s+", " ", ocr(crop) if crop.size[0] > side else "").strip(" |[]_-—")
            marks.append((b["checked"], label or "(без подписи)"))
    return marks


def marks_section(marks: list, page_no: Optional[int] = None) -> str:
    """Секция в формате экстрактора .docx — её читает fields.checkbox_marks."""
    if not marks:
        return ""
    head = "=== ОТМЕТКИ (распознаны на скане" + (f", стр. {page_no}" if page_no else "") + ") ==="
    lines = [head] + [f"{'☒ ОТМЕЧЕНО' if c else '☐ НЕ ОТМЕЧЕНО'} | {label}" for c, label in marks]
    return "\n".join(lines)
