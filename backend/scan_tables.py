"""Текст страницы скана с таблицами по ячейкам.

Обычный вывод Tesseract читает таблицу сплошняком и перемешивает слова соседних колонок:
«совещание руководства по предварительног / 14:30 качеству, о совещания». Поэтому:
1) на картинке ищутся линии таблиц, по ним — ячейки;
2) Tesseract отдаёт слова с координатами (image_to_data) — один вызов на страницу;
3) слова раскладываются по ячейкам, строка таблицы выводится как «ячейка | ячейка | …» —
   так же, как таблицы .docx, на которые рассчитаны правила проверки.

Только Pillow, без numpy.
"""
import re
from typing import Optional

from PIL import Image

from scan_marks import _binary_rows

_GAP = 3  # разрыв линии в пикселях, который считается сканерным шумом


def _long_runs(row: bytes, min_len: int, gap: int = _GAP) -> list:
    """Отрезки чернил длиной от min_len с разрывами до gap пикселей."""
    out = []
    for m in re.finditer(rb"\x00(?:\xff{0,%d}\x00)*" % gap, row):
        if m.end() - m.start() >= min_len:
            out.append((m.start(), m.end()))
    return out


def _merge_lines(raw: list, tol: int = 3) -> list:
    """Соседние по высоте отрезки (толщина линии) → одна линия (y, x0, x1)."""
    lines = []
    for y, x0, x1 in sorted(raw):
        for ln in lines:
            if 0 <= y - ln["y1"] <= tol and min(x1, ln["x1"]) - max(x0, ln["x0"]) > 0.5 * min(x1 - x0, ln["x1"] - ln["x0"]):
                ln["y1"] = y
                ln["x0"], ln["x1"] = min(ln["x0"], x0), max(ln["x1"], x1)
                break
        else:
            lines.append({"y0": y, "y1": y, "x0": x0, "x1": x1})
    return [((ln["y0"] + ln["y1"]) // 2, ln["x0"], ln["x1"]) for ln in lines]


def _lines(img: Image.Image, dpi: int) -> tuple:
    """(горизонтали [(y, x0, x1)], вертикали [(x, y0, y1)]) — линии таблиц и рамок полей."""
    rows = _binary_rows(img)
    w = img.size[0]
    hlines = _merge_lines([(y, a, b) for y, r in enumerate(rows) for a, b in _long_runs(r, int(dpi * 0.35))])
    cols = _binary_rows(img.transpose(Image.Transpose.ROTATE_90))
    # ROTATE_90 — против часовой: столбец x исходной картинки стал строкой w-1-x
    raw = [(w - 1 - i, a, b) for i, c in enumerate(cols) for a, b in _long_runs(c, int(dpi * 0.2))]
    vlines = [(x, a, b) for x, a, b in _merge_lines(raw)]
    return hlines, vlines


def _tables(hlines: list, vlines: list, tol: int = 6) -> list:
    """Таблицы — связные группы пересекающихся линий: [(горизонтали, вертикали)]."""
    items = [("h", ln) for ln in hlines] + [("v", ln) for ln in vlines]
    parent = list(range(len(items)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i, (ki, (yi, a0, a1)) in enumerate(items):
        if ki != "h":
            continue
        for j, (kj, (xj, b0, b1)) in enumerate(items):
            if kj == "v" and a0 - tol <= xj <= a1 + tol and b0 - tol <= yi <= b1 + tol:
                parent[find(i)] = find(j)
    groups = {}
    for i, it in enumerate(items):
        groups.setdefault(find(i), []).append(it)
    out = []
    for g in groups.values():
        hs = [ln for k, ln in g if k == "h"]
        vs = [ln for k, ln in g if k == "v"]
        if len(hs) >= 2 and len(vs) >= 2:
            out.append((hs, vs))
    return out


def _dedupe(values: list, tol: int) -> list:
    out = []
    for v in sorted(values):
        if out and v - out[-1] <= tol:
            continue
        out.append(v)
    return out


def find_cells(img: Image.Image, dpi: int) -> list:
    """Строки таблиц на странице: [{"y0", "y1", "cells": [(x0, x1), …]}] сверху вниз.

    Границы строк — горизонтали, пересекающие левую колонку таблицы (в Плане это колонка
    времени): объединённые ячейки правее не дробят строку. Границы колонок — вертикали,
    проходящие через середину строки."""
    hlines, vlines = _lines(img, dpi)
    bands = []
    for hs, vs in _tables(hlines, vlines):
        xs = _dedupe([x for x, _, _ in vs], 8)
        if len(xs) < 2:
            continue
        left, right = xs[0], xs[-1]
        anchor = (xs[0] + xs[1]) / 2
        ys = _dedupe([y for y, x0, x1 in hs if x0 - 6 <= anchor <= x1 + 6], 8)
        for y0, y1 in zip(ys, ys[1:]):
            if y1 - y0 < dpi * 0.12:
                continue
            mid = (y0 + y1) / 2
            cuts = _dedupe([x for x, a, b in vs if a - 6 <= mid <= b + 6], 8)
            cuts = sorted(set(cuts) | {left, right})
            cells = [(a, b) for a, b in zip(cuts, cuts[1:]) if b - a >= 15]
            if cells:
                bands.append({"y0": y0, "y1": y1, "cells": cells})
    # строки разных таблиц на одной высоте (поля шапки Листа оценки) — одна строка
    out = []
    for b in sorted(bands, key=lambda b: b["y0"]):
        for o in out:
            inter = min(o["y1"], b["y1"]) - max(o["y0"], b["y0"])
            if inter > 0.6 * min(o["y1"] - o["y0"], b["y1"] - b["y0"]):
                o["y0"], o["y1"] = min(o["y0"], b["y0"]), max(o["y1"], b["y1"])
                o["cells"] = sorted(set(o["cells"]) | set(b["cells"]))
                break
        else:
            out.append(dict(b))
    return out


def words_from_data(data: dict, min_conf: float = 0) -> list:
    """Слова из pytesseract.image_to_data(output_type=DICT): [{"text", "x0", "y0", "x1", "y1", "line"}]."""
    words = []
    for i, text in enumerate(data.get("text", [])):
        text = (text or "").strip()
        try:
            conf = float(data["conf"][i])
        except (TypeError, ValueError):
            conf = -1
        if not text or conf < min_conf:
            continue
        x, y, w, h = (int(data[k][i]) for k in ("left", "top", "width", "height"))
        words.append({"text": text, "x0": x, "y0": y, "x1": x + w, "y1": y + h,
                      "line": (data["block_num"][i], data["par_num"][i], data["line_num"][i])})
    return words


def _join_words(words: list) -> str:
    """Слова одного куска в порядке чтения: по строкам Tesseract, внутри строки — слева направо."""
    lines = {}
    for wd in words:
        lines.setdefault(wd["line"], []).append(wd)
    ordered = sorted(lines.values(), key=lambda ws: sum(w["y0"] for w in ws) / len(ws))
    return " ".join(" ".join(w["text"] for w in sorted(ws, key=lambda w: w["x0"])) for ws in ordered)


def page_text(words: list, table_rows: list, fill_empty=None) -> str:
    """Текст страницы: строки вне таблиц — как у Tesseract, строки таблиц — «ячейка | ячейка».
    Слова на высоте строки таблицы, но вне её ячеек (подпись слева от поля формы, «7» в
    клетке и «Количество …» справа от неё) становятся отдельными ячейками той же строки.
    fill_empty((x0, y0, x1, y1)) — текст ячейки, где Tesseract на всей странице ничего не
    нашёл: одиночную цифру номера строки он отбрасывает как шум."""
    rows = [{"y0": r["y0"], "y1": r["y1"], "cells": r["cells"], "words": {}, "free": []} for r in table_rows]
    free = []
    for wd in words:
        cx, cy = (wd["x0"] + wd["x1"]) / 2, (wd["y0"] + wd["y1"]) / 2
        row = next((r for r in rows if r["y0"] < cy < r["y1"]), None)
        if row is None:
            free.append(wd)
            continue
        k = next((k for k, (a, b) in enumerate(row["cells"]) if a < cx < b), None)
        if k is None:
            row["free"].append(wd)
        else:
            row["words"].setdefault(k, []).append(wd)

    items = []  # (y, текст)
    for r in rows:
        parts = []
        for k, (a, b) in enumerate(r["cells"]):
            text = _join_words(r["words"].get(k, []))
            if not text and fill_empty:
                text = fill_empty((a, r["y0"], b, r["y1"])) or ""
            parts.append((a, text))
        # слова вне ячеек: группы между ячейками
        bounds = [a for a, _ in r["cells"]] + [float("inf")]
        groups = {}
        for wd in r["free"]:
            cx = (wd["x0"] + wd["x1"]) / 2
            groups.setdefault(next(i for i, a in enumerate(bounds) if cx < a), []).append(wd)
        for g in groups.values():
            parts.append((min(w["x0"] for w in g), _join_words(g)))
        cells = [t for _, t in sorted(parts, key=lambda p: p[0])]
        if any(cells):
            items.append((r["y0"], " | ".join(cells)))
    for ws in _visual_lines(free):
        items.append((min(w["y0"] for w in ws), " ".join(w["text"] for w in ws)))
    return "\n".join(text for _, text in sorted(items, key=lambda it: it[0]))


def _visual_lines(words: list) -> list:
    """Слова вне таблиц — по строкам на странице: одна строка — слова, чьи середины по высоте
    рядом. Tesseract иногда разносит по разным строкам подпись поля и значение на линии
    («7.1 Руководитель экспертной группы» и «Сукорская О.А.»), и значение уходит строкой выше."""
    lines = []
    for wd in sorted(words, key=lambda w: (w["y0"] + w["y1"]) / 2):
        cy, h = (wd["y0"] + wd["y1"]) / 2, wd["y1"] - wd["y0"]
        if lines and abs(cy - lines[-1]["cy"]) <= 0.5 * max(h, lines[-1]["h"]):
            ln = lines[-1]
            ln["words"].append(wd)
            ln["cy"] = sum((w["y0"] + w["y1"]) / 2 for w in ln["words"]) / len(ln["words"])
        else:
            lines.append({"cy": cy, "h": h, "words": [wd]})
    return [sorted(ln["words"], key=lambda w: w["x0"]) for ln in lines]


def erase_lines(img: Image.Image, dpi: int) -> Image.Image:
    """Копия страницы без линий таблиц и рамок: иначе Tesseract читает их как «|» и склеивает
    слова соседних ячеек («16:30-|Анализ»). Квадраты отметок короче порогов и остаются.
    Линия здесь — без разрывов длиннее пикселя: с допуском в 3 пикселя нижние края букв
    одного слова сливаются в «линию», и стираются хвосты «ц» и «щ»."""
    from PIL import ImageDraw

    out = img.convert("L").copy()
    draw = ImageDraw.Draw(out)
    rows = _binary_rows(out)
    for y, r in enumerate(rows):
        for a, b in _long_runs(r, int(dpi * 0.35), gap=1):
            draw.rectangle((a, y - 1, b, y + 1), fill=255)
    w = out.size[0]
    cols = _binary_rows(img.convert("L").transpose(Image.Transpose.ROTATE_90))
    for i, c in enumerate(cols):
        x = w - 1 - i
        for a, b in _long_runs(c, int(dpi * 0.2), gap=1):
            draw.rectangle((x - 1, a, x + 1, b), fill=255)
    return out


def has_ink(img: Image.Image, box: tuple, inset: int = 5, min_pixels: int = 30) -> bool:
    """Есть ли чернила внутри прямоугольника (без краёв, где остатки линий). Порог — число
    пикселей, а не доля: одиночная «1» в высокой ячейке — сотая доля площади."""
    x0, y0, x1, y1 = box
    crop = img.crop((x0 + inset, y0 + inset, x1 - inset, y1 - inset)).convert("L")
    if crop.size[0] <= 0 or crop.size[1] <= 0:
        return False
    return crop.point(lambda v: 0 if v < 160 else 255).tobytes().count(0) >= min_pixels


def skew_angle(img: Image.Image, max_deg: float = 3.0) -> float:
    """Наклон скана в градусах (для Image.rotate): угол, при котором строки текста и линии
    таблиц горизонтальны — у проекции чернил на вертикаль самые резкие пики. Сначала грубо
    (шаг 0,5°), потом точнее (0,05°) на уменьшенной копии."""
    small = img.convert("L")
    k = 800 / small.size[0]
    small = small.resize((800, int(small.size[1] * k))).point(lambda v: 0 if v < 160 else 255)
    w = small.size[0]

    def score(angle: float) -> int:
        data = small.rotate(angle, fillcolor=255).tobytes()
        return sum(data[y * w:(y + 1) * w].count(0) ** 2 for y in range(len(data) // w))

    best = max((a / 2 for a in range(int(-max_deg * 2), int(max_deg * 2) + 1)), key=score)
    fine = [best + a / 20 for a in range(-10, 11)]
    return round(max(fine, key=score), 2)


def deskew(img: Image.Image) -> Image.Image:
    angle = skew_angle(img)
    if abs(angle) < 0.05:
        return img
    return img.rotate(angle, resample=Image.Resampling.BICUBIC, fillcolor="white")


def upright_angle(osd: Optional[str]) -> int:
    """Угол из вывода pytesseract.image_to_osd: «Rotate: 90» — на сколько градусов по часовой
    повернуть картинку, чтобы текст встал прямо (0 — не поворачивать)."""
    m = re.search(r"Rotate:\s*(\d+)", osd or "")
    return int(m.group(1)) % 360 if m else 0
