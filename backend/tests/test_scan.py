"""Распознавание сканов без Tesseract: квадраты-отметки, ячейки таблиц, сборка текста страницы
по словам с координатами, наклон и поворот. Картинки рисуются в тесте."""
from PIL import Image, ImageDraw

import scan_marks
import scan_tables

DPI = 200


def page(w=1600, h=1000):
    img = Image.new("L", (w, h), 255)
    return img, ImageDraw.Draw(img)


def box(draw, x, y, side=28, checked=False):
    draw.rectangle((x, y, x + side, y + side), outline=0, width=2)
    if checked:
        draw.line((x + 4, y + 4, x + side - 4, y + side - 4), fill=0, width=3)
        draw.line((x + 4, y + side - 4, x + side - 4, y + 4), fill=0, width=3)


def test_boxes_checked_and_empty():
    img, d = page()
    box(d, 160, 100)
    box(d, 840, 100)
    box(d, 160, 180, checked=True)                  # ☒ Второй инспекционный контроль
    found = scan_marks.find_boxes(img, DPI)
    assert [b["checked"] for b in found] == [False, False, True]
    assert [b["line"] for b in found] == [0, 0, 1]


def test_box_touching_letters_and_solid_block_are_not_marks():
    img, d = page()
    d.rectangle((150, 100, 178, 128), outline=0, width=2)
    d.rectangle((130, 105, 147, 125), fill=0)       # соседняя буква вплотную слева
    d.rectangle((400, 100, 428, 128), fill=0)       # сплошной чёрный прямоугольник
    assert scan_marks.find_boxes(img, DPI) == []


def test_grid_rows_get_common_label():
    img, d = page()
    for x in (700, 850, 1000, 1150, 1300):
        box(d, x, 300, checked=(x == 700))
    marks = scan_marks.label_marks(img, scan_marks.find_boxes(img, DPI), ocr=lambda im: "не вызывается")
    assert marks == [(True, "строка таблицы оценок 1")] + [(False, "строка таблицы оценок 1")] * 4


def test_labels_from_ocr_and_section_format():
    img, d = page()
    box(d, 160, 100, checked=True)
    box(d, 840, 100)
    labels = iter(["Второй инспекционный контроль", "Расширение области сертификации"])
    marks = scan_marks.label_marks(img, scan_marks.find_boxes(img, DPI), ocr=lambda im: next(labels))
    assert marks == [(True, "Второй инспекционный контроль"), (False, "Расширение области сертификации")]
    section = scan_marks.marks_section(marks)
    assert section.splitlines()[1] == "☒ ОТМЕЧЕНО | Второй инспекционный контроль"
    assert section.splitlines()[2] == "☐ НЕ ОТМЕЧЕНО | Расширение области сертификации"


def table(draw, x0, y0, cols, rows):
    """Сетка: cols — x границ колонок, rows — y границ строк."""
    for y in rows:
        draw.line((cols[0], y0 + y, cols[-1], y0 + y), fill=0, width=2)
    for x in cols:
        draw.line((x, y0 + rows[0], x, y0 + rows[-1]), fill=0, width=2)


def test_cells_and_merged_cells():
    img, d = page()
    table(d, 0, 100, [100, 250, 700, 1200], [0, 120, 240])
    # объединённая ячейка во всю ширину («2-й день …»)
    d.line((100, 440, 1200, 440), fill=0, width=2)
    d.line((100, 340, 100, 440), fill=0, width=2)
    d.line((1200, 340, 1200, 440), fill=0, width=2)
    rows = scan_tables.find_cells(img, DPI)
    assert [len(r["cells"]) for r in rows] == [3, 3, 1]


def word(text, x0, y0, line, w=80, h=24):
    return {"text": text, "x0": x0, "y0": y0, "x1": x0 + w, "y1": y0 + h, "line": line}


def test_page_text_puts_words_into_cells():
    rows = [{"y0": 100, "y1": 220, "cells": [(100, 250), (250, 700), (700, 1200)]}]
    words = [
        word("6", 300, 40, (1, 1, 1)), word("Сроки", 120, 40, (1, 1, 1)),
        # одна строка Tesseract, но слова в разных колонках таблицы
        word("16:30-", 110, 110, (2, 1, 1)), word("Анализ", 260, 110, (2, 1, 1)), word("6.1.1,", 720, 110, (2, 1, 1)),
        word("17:00", 110, 140, (2, 1, 2)), word("рисков", 260, 140, (2, 1, 2)),
        word("Итог", 120, 300, (3, 1, 1)),
    ]
    assert scan_tables.page_text(words, rows).splitlines() == [
        "Сроки 6",
        "16:30- 17:00 | Анализ рисков | 6.1.1,",
        "Итог",
    ]


def test_page_text_words_beside_box_join_row():
    """п.24 Акта: число в клетке, подпись справа — одна строка «7 | Количество …»."""
    rows = [{"y0": 100, "y1": 160, "cells": [(100, 190)]}]
    words = [word("7", 130, 115, (1, 1, 1), w=20), word("Количество", 220, 115, (1, 1, 2)),
             word("малозначительных", 340, 115, (1, 1, 2))]
    assert scan_tables.page_text(words, rows) == "7 | Количество малозначительных"


def test_words_from_tesseract_data_skip_empty_and_negative_conf():
    data = {"text": ["", "Акт", "№", " "], "conf": ["-1", "96.5", "88", "-1"],
            "left": [0, 10, 60, 0], "top": [0, 5, 5, 0], "width": [0, 40, 10, 0], "height": [0, 20, 20, 0],
            "block_num": [1, 1, 1, 1], "par_num": [1, 1, 1, 1], "line_num": [1, 1, 1, 1]}
    assert [w["text"] for w in scan_tables.words_from_data(data)] == ["Акт", "№"]


def test_skew_is_detected_and_removed():
    img, d = page(1000, 800)
    for y in range(100, 700, 60):
        d.rectangle((100, y, 900, y + 14), fill=0)   # строки «текста»
    tilted = img.rotate(1.2, fillcolor=255)
    assert abs(scan_tables.skew_angle(tilted) + 1.2) <= 0.1


def test_upright_angle_from_osd():
    osd = "Page number: 0\nOrientation in degrees: 270\nRotate: 90\nOrientation confidence: 4.1\n"
    assert scan_tables.upright_angle(osd) == 90
    assert scan_tables.upright_angle("") == 0
