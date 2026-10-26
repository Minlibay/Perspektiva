import checklists
import doc_types
import formula_checks as FC
import stage3_docs as D

O1_TPL = "Чек лист  отчет —1 этап.docx"
Z1_TPL = "Чек лист Заключение —1 этап.docx"


def run(checklist_path, tpl, ctype, checked: dict, sources: dict):
    items = checklists.parse_checklist(str(checklist_path(tpl)))
    all_texts = {**checked, **sources}
    res = FC.run_checklist(ctype, items, checked, all_texts, doc_types.classify_all(all_texts))
    return {it["item_no"]: ("ok" if v["ok"] else "nok" if v["nok"] else "manual", v["reason"])
            for it, v in zip(items, res["checklist"])}


def o1(checklist_path, sources=None, **kw):
    return run(checklist_path, O1_TPL, "otchet1", {"Отчет 1 этапа.docx": D.otchet1(**kw)},
               D.SOURCES if sources is None else sources)


def z1(checklist_path, o1_kw=None, sources=None, **kw):
    src = dict(D.SOURCES if sources is None else sources)
    src["Отчет 1 этапа.docx"] = D.otchet1(**(o1_kw or {}))
    return run(checklist_path, Z1_TPL, "zakl1", {"Заключение 1 этапа.docx": D.zakl1(**kw)}, src)


def statuses(r):
    return [s for s, _ in r.values()]


def test_every_item_has_rule(checklist_path):
    for tpl, ctype in ((O1_TPL, "otchet1"), (Z1_TPL, "zakl1")):
        items = checklists.parse_checklist(str(checklist_path(tpl)))
        assert all(FC.rule_for(ctype, it) for it in items), tpl


def test_otchet1_all_ok(checklist_path):
    r = o1(checklist_path)
    assert statuses(r) == ["ok"] * 7, r


def test_otchet1_short_name_and_late_approval(checklist_path):
    # образец УТЕКО: в «наименование … полностью» стоит «ООО «УТЕКО»» — сокращённое
    r = o1(checklist_path, org="ООО «Ромашка»", approval="«06» июля 2026г.")
    assert r[1][0] == "nok" and "сокращённое" in r[1][1]
    assert r[2][0] == "nok" and "03.07.2026" in r[2][1]      # 26.06 + 5 р.д.


def test_otchet1_basis_without_dop_agreement(checklist_path):
    r = o1(checklist_path, basis="Договор № 28/01-2026 от 28 января 2026 г.")
    assert r[4][0] == "nok" and "10.03.2026" in r[4][1]      # ДС № 1 загружен, но не указан
    # ДС указан, но не загружен — на ручную, а не NOK
    src = {k: v for k, v in D.SOURCES.items() if k != "Доп соглашение 1.pdf"}
    r = o1(checklist_path, sources=src)
    assert r[4][0] == "manual" and "не загружено доп. соглашение" in r[4][1]


def test_otchet1_team_and_nc(checklist_path):
    r = o1(checklist_path, expert="К.К. Кузнецов", nc_count=3,
           conformity="Документация СМК соответствует процессам и деятельности организации")
    assert r[5][0] == "nok" and "Кузнецов К.К." in r[5][1]
    assert r[6][0] == "nok"                                     # 3 ≠ 2 строки таблицы
    assert r[7][0] == "nok" and "в целом" in r[7][1]


def test_otchet1_no_nc(checklist_path):
    r = o1(checklist_path, nc_count=0, rows=0,
           conformity="Документация СМК соответствует процессам и деятельности организации")
    assert r[6][0] == "ok" and r[7][0] == "ok", r


def test_zakl1_all_ok(checklist_path):
    r = z1(checklist_path)
    assert statuses(r) == ["ok"] * 7, r


def test_zakl1_problems(checklist_path):
    r = z1(checklist_path, approver="И.И. Иванова", approval="«23» июля 2026 г.", sign_date="23.07.2026",
           rows=1, eliminated=False, second_stage=False, dist_org="ООО «Лютик»")
    assert r[2][0] == "nok"                                     # утвердил сам РЭГ
    assert r[3][0] == "nok" and "21.07.2026" in r[3][1]         # 14.07 + 5 р.д.
    assert r[4][0] == "nok"                                     # рассылка — другая организация
    assert r[5][0] == "nok"                                     # 1 строка, в Отчете 2 НС
    assert r[6][0] == "nok" and r[7][0] == "nok"


def test_zakl1_missing_sources(checklist_path):
    r = run(checklist_path, Z1_TPL, "zakl1", {"Заключение 1 этапа.docx": D.zakl1()}, {})
    assert r[1][0] == "nok" and "Отчет по результатам 1 этапа" in r[1][1]
    assert r[3][0] == "nok" and "Отчет об устранении" in r[3][1]


def test_uteko_sample(samples_dir, checklist_path):
    """Отчет 1 этапа УТЕКО. Договор УТЕКО — скан (без OCR не читается); Приказа и Разбивки
    УТЕКО в образцах нет. Количество НС (8) и п.10.1.4 сверяются внутри Отчета."""
    import main

    text = main.extract_text_from_docx(str(samples_dir / "ИИ Отчет 1 этап РА УТЕКО .docx"))
    r = run(checklist_path, O1_TPL, "otchet1", {"Отчет.docx": text}, {})
    assert r[2][0] == "ok" and "01.07.2026" in r[2][1]
    assert r[6][0] == "ok" and "8" in r[6][1]
    assert r[7][0] == "ok"
    # реквизиты основания — с OCR-текстом Договора (на сервере его даёт Tesseract)
    ocr = "ДОГОВОР № 28/01-01-2026\nг. Москва 28 января 2026 г.\nОбщество с ограниченной " \
          "ответственностью «УТЕКО» (ООО «УТЕКО»), именуемое в дальнейшем «Заказчик»"
    r = run(checklist_path, O1_TPL, "otchet1", {"Отчет.docx": text}, {"Договор.pdf": ocr})
    assert r[4][0] == "ok", r[4]
    assert r[1][0] == "nok" and "сокращённое" in r[1][1]       # «ООО «УТЕКО»» вместо полного
