import json
from datetime import date

import pytest

import checklists
import main


def test_plan_2026_items_bound_to_rules(checklist_path):
    items = checklists.parse_checklist(str(checklist_path("ИИ -План АУДИТА.docx")))
    assert [main._plan_rule_id(it) for it in items] == [18, 3, 4, 5, 6, 8, 19, 16, 12, 17, 14, 15]


@pytest.mark.parametrize("area,comments,rid", [
    # пункты редакции 2025 (16 пунктов) — по тексту, а не по позиции
    ("Дата утверждения/согласования в шапке", "ИИ1 не позднее 5 рабочих дней до начала аудита ИИ2", 0),
    ("Соответствие организационной структуры с подразделениями, указанными в плане", "", 1),
    ("Трудоемкость", "", 2),
    ("п.5 Сроки проведения аудита", "", 7),
    ("п. Исключения", "Только 7.1.3.5 и 7.1.4.3", 10),
    ("Предварительное/\nзаключительное совещание на каждой площадке", "", 11),
    ("Наличие всех процессов с указанием видов деятельности", "", 13),
    ("Наличие пунктов 4.4.1, 4.4.2, 4.4.3 там, где выделены процессы", "", 14),
    ("Что-то новое", "", None),
])
def test_plan_legacy_items_bound_by_text(area, comments, rid):
    assert main._plan_rule_id({"area": area, "comments": comments}) == rid


def test_svod_2026_items_bound_to_rules(checklist_path):
    items = checklists.parse_checklist(str(checklist_path("ЧК -Сводный акт.docx")))
    # №6, 9, 12 исключены; №13 теперь «Аутсорсинг» (ключ 14)
    assert [main._svod_rule_id(it) for it in items] == [0, 1, 2, 3, 4, 6, 7, 9, 10, 14, 13]


def test_add_workdays_matches_template_example():
    # Шаблон: «если начало 10.02.2026, то утверждение не позднее 03.02.2026»
    assert main._add_workdays(date(2026, 2, 10), -5) == date(2026, 2, 3)
    assert main._add_workdays(date(2026, 7, 27), -5) == date(2026, 7, 20)
    assert main._add_workdays(date(2026, 7, 17), 5) == date(2026, 7, 24)


def test_plan_and_prikaz_dates():
    plan = "УТВЕРЖДАЮ\n«17» июля 2026  г.\n5 Сроки проведения аудита | 27-30.07.2026\n"
    assert main._plan_approval_date(plan) == date(2026, 7, 17)
    assert main._plan_start_date(plan) == date(2026, 7, 27)
    assert main._prikaz_date("П Р И К А З\n№ П/13-07/01 13 июля 2026 г.\nО назначении") == date(2026, 7, 13)
    assert main._prikaz_date("ПРИКАЗ № 5 от 03.07.2026") == date(2026, 7, 3)


def test_exclusion_clauses():
    plan = ("8 Исключения из требований стандартов | п. 7.1.3.5 (абзац 1), п. 7.1.4.3 (абзац 1), п.8.3.5.4\n"
            "СТО Газпром 9001-2018\n9 Рабочий язык проведения аудита | Русский\n4.4.1, 4.4.2")
    cert = "Исключения из требований стандарта СТО Газпром 9001-2018: п. 7.1.3.5 (абзац 1), п. 7.1.4.3 (абзац 1)"
    assert main._exclusion_clauses(main._exclusion_block(plan)) == ["7.1.3.5", "7.1.4.3", "8.3.5.4"]
    assert main._exclusion_clauses(main._exclusion_block(cert)) == ["7.1.3.5", "7.1.4.3"]
    assert main._exclusion_block("нет раздела") is None


SCHEDULE = """10 Объекты аудита
1-й день. 27.07.2026
(г. Пермь, ул. Ленина, д. 1)
10:00 | Предварительное совещание | ...
18:00 | Рабочее совещание
2-й день. 28.07.2026
(г. Пермь, ул. Ленина, д. 1)
12:00 | Заключительное совещание
3-й день. 29.07.2026
(г. Кунгур, ул. Заводская, д. 5)
10:00 | Инструктаж
4-й день. 30.07.2026
12:00 | Заключительное совещание
=== ОТМЕТКИ (галочки/крестики) В ДОКУМЕНТЕ ===
☐ Предварительное совещание
"""


def test_schedule_days_and_addresses():
    days = main._plan_schedule_days(SCHEDULE)
    assert [d["address"] for d in days] == [
        "г. Пермь, ул. Ленина, д. 1", "г. Пермь, ул. Ленина, д. 1",
        "г. Кунгур, ул. Заводская, д. 5", "г. Кунгур, ул. Заводская, д. 5",  # день без адреса — прежний
    ]
    assert "ОТМЕТКИ" not in days[-1]["text"]


def test_meetings_by_site():
    final = main._plan_meeting_by_site(SCHEDULE, "заключительное совещан", "last")
    assert final["sites"] == 2 and final["missing"] == []
    opening = main._plan_meeting_by_site(SCHEDULE, "предварительное совещан", "first")
    assert opening["missing"] == ["г. Кунгур, ул. Заводская, д. 5, 3-й день. 29.07.2026"]
    assert main._plan_meeting_by_site("без графика", "заключительное совещан", "last") is None


@pytest.fixture
def stub_gigachat(monkeypatch):
    """Модель всегда отвечает NOK — всё прочее в результате поставили детерминированные проверки."""
    stub = json.dumps({"ok": False, "nok": True, "reason": "MODEL-STUB"}, ensure_ascii=False)
    monkeypatch.setattr(main, "_gigachat_call", lambda *a, **k: stub)


def _plan_text(items_needed=()):
    return (
        "ПЛАН аудита системы менеджмента качества\nУТВЕРЖДАЮ\n«17» июля 2026 г.\n"
        "1 Цель и область аудита\n5 Сроки проведения аудита | 27-30.07.2026\n"
        "8 Исключения из требований стандартов | п. 7.1.3.5 (абзац 1), п. 7.1.4.3 (абзац 1)\n"
        "СТО Газпром 9001-2018\n9 Рабочий язык проведения аудита | Русский\n" + SCHEDULE
    )


def _run_plan(checklist_path, all_texts, header_type="Второй инспекционный контроль", monkeypatch=None):
    monkeypatch.setattr(main, "extract_header_info", lambda *a, **k: {
        "Наименование Заявителя": "ООО «Ромашка»", "Вид аудита": header_type,
        "Даты проведения": "27-30.07.2026", "РЭГ": "—",
    })
    items = checklists.parse_checklist(str(checklist_path("ИИ -План АУДИТА.docx")))
    plan = _plan_text()
    res = main.process_checklist_advanced("key", {"План.docx": plan, **all_texts}, items, {},
                                          "GigaChat", plan)
    return {main._plan_rule_id(it): v for it, v in zip(items, res["checklist"])}


def test_plan_2026_deterministic_rules(checklist_path, stub_gigachat, monkeypatch):
    prikaz = "ПРИКАЗ\n№ П/13-07/01 13 июля 2026 г.\nО назначении экспертной группы"
    cert = ("ПРИЛОЖЕНИЕ к сертификату соответствия\nИсключения из требований стандарта "
            "СТО Газпром 9001-2018: п. 7.1.3.5 (абзац 1), п. 7.1.4.3 (абзац 1)")
    by_rid = _run_plan(checklist_path, {"Приказ.pdf": prikaz, "скан.pdf": cert}, monkeypatch=monkeypatch)

    # п.1: 17.07 ≤ 27.07 − 5 р.д. (20.07) и не раньше Приказа 13.07 — прежний код давал NOK
    assert by_rid[18]["ok"], by_rid[18]["reason"]
    assert "D_приказа=13.07.2026" in by_rid[18]["reason"]
    # п.7: ИК — сверка с Сертификатом (найден по содержимому, имя «скан.pdf»)
    assert by_rid[19]["ok"], by_rid[19]["reason"]
    assert "скан.pdf" in by_rid[19]["reason"]
    # п.8: заключительное на обоих адресах; п.10: на втором адресе нет предварительного
    assert by_rid[16]["ok"], by_rid[16]["reason"]
    assert by_rid[17]["nok"] and "Кунгур" in by_rid[17]["reason"]


def test_plan_approval_too_late_and_before_prikaz(checklist_path, stub_gigachat, monkeypatch):
    prikaz = "ПРИКАЗ\n№ 7 20 июля 2026 г.\nО назначении экспертной группы"
    by_rid = _run_plan(checklist_path, {"Приказ.pdf": prikaz}, monkeypatch=monkeypatch)
    assert by_rid[18]["nok"]
    assert "раньше Приказа об ЭГ от 20.07.2026" in by_rid[18]["reason"]


def test_plan_exclusions_mismatch_and_fallback(checklist_path, stub_gigachat, monkeypatch):
    otchet1 = ("Отчет по результатам первого этапа ресертификационного аудита\n"
               "8 Исключения из требований стандартов | п. 7.1.3.5 (абзац 1)\n9 Рабочий язык")
    by_rid = _run_plan(checklist_path, {"Отчет 1 этап.docx": otchet1},
                       header_type="Ресертификационный аудит", monkeypatch=monkeypatch)
    assert by_rid[19]["nok"]
    assert "только в Плане: 7.1.4.3" in by_rid[19]["reason"]

    # Источника нет — прежнее правило допустимых исключений, с пометкой
    by_rid = _run_plan(checklist_path, {}, monkeypatch=monkeypatch)
    assert by_rid[19]["ok"]
    assert by_rid[19]["reason"].startswith("Источник для сверки исключений")


def test_plan_dki_sample(samples_dir, checklist_path, stub_gigachat):
    """Пакет ДКИ (2-й ИК): План утверждён 17.07, Приказ 13.07, начало 27.07 —
    прежнее правило «окна ≤5 рабочих дней» давало здесь NOK."""
    names = ["ИИ План аудита.docm", "ИИ Приказ об ЭГ.pdf", "Информация для подготовки к аудиту (2).xlsx",
             "ИИ Разбивка области сертификации по кодам ОКВЭД ДКИ (1).docx"]
    texts = {}
    for n in names:
        p = samples_dir / n
        if p.suffix.lower() in (".docx", ".docm"):
            texts[n] = main.extract_text_from_docx(str(p))
        elif p.suffix.lower() == ".pdf":
            texts[n] = main.extract_text_from_pdf(str(p))
        else:
            texts[n] = main.extract_text_from_xlsx(str(p))
    items = checklists.parse_checklist(str(checklist_path("ИИ -План АУДИТА.docx")))
    res = main.process_checklist_advanced("key", texts, items, {}, "GigaChat", texts[names[0]])
    by_rid = {main._plan_rule_id(it): v for it, v in zip(items, res["checklist"])}

    assert by_rid[18]["ok"], by_rid[18]["reason"]
    assert "D_утв=17.07.2026, D_начало=27.07.2026, D_приказа=13.07.2026" in by_rid[18]["reason"]
    assert by_rid[3]["ok"], by_rid[3]["reason"]   # вид аудита: «второй инспекционный»
    assert by_rid[16]["ok"] and by_rid[17]["ok"]


def test_fill_plan_2026(checklist_path, tmp_path):
    items = checklists.parse_checklist(str(checklist_path("ИИ -План АУДИТА.docx")))
    data = {
        "header": {"Наименование Заявителя": "ООО «Ромашка»", "Вид аудита": "ИК2",
                   "Даты проведения": "27-30.07.2026", "РЭГ": "ОГН0.RU.0136"},
        "checklist": [{"ok": True, "nok": False}] * 6 + [{"ok": False, "nok": True}] * 6,
    }
    out = tmp_path / "out.docx"
    main.fill_plan_with_checklist(str(checklist_path("ИИ -План АУДИТА.docx")), data, str(out), items)

    marks = checklists.read_marks(str(out), ok_marks=("☑",), nok_marks=("☒",))
    assert [m["ok"] for m in marks] == [True] * 6 + [False] * 6
    assert [m["nok"] for m in marks] == [False] * 6 + [True] * 6
    header = checklists.open_docx(out).tables[0]
    assert header.rows[0].cells[1].text == "ООО «Ромашка»"
    assert header.rows[3].cells[1].text == "ОГН0.RU.0136"


def test_svod_outsourcing_must_be_filled():
    item = {"area": "Аутсорсинг", "comments": "СВ7=А50"}
    empty = ("СВОДНЫЙ АКТ\n6. Организация передает на аутсорсинг следующие процессы:\n"
             "7. Для изготовления продукции организация закупает")
    filled = empty.replace("процессы:\n", "процессы:\nМеталлообработка — ООО «Цех»\n")
    v = main._svod_deterministic_postcheck(14, {"ok": True, "nok": False, "reason": ""}, empty, {}, item)
    assert v["nok"]
    v = main._svod_deterministic_postcheck(14, {"ok": True, "nok": False, "reason": "m"}, filled, {}, item)
    assert v["ok"] and v["reason"] == "m"


def test_svod_headcount_source_follows_template():
    svod = "СВОДНЫЙ АКТ\nОбщие сведения об организации\nЧисленность персонала в системе менеджмента | 360"
    trud = {"Расчет трудоемкости.docx": "РАСЧЕТ ТРУДОЕМКОСТИ\nЧисленность персонала, на деятельность "
                                        "которого распространяется СМК | 360\nКоличество смен | 1"}
    item_2026 = {"area": "Численность сотрудников", "comments": "СВ4= ТР1"}
    v = main._svod_deterministic_postcheck(7, {"ok": False, "nok": True, "reason": ""}, svod, trud, item_2026)
    assert v["ok"] and "Расчёт трудоёмкости — 360" in v["reason"]
