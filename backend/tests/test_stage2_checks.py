import checklists
import doc_types
import formula_checks as FC
import stage2_docs as D

AKT_TPL = "Чек лист Акт.docx"
ZAKL_TPL = "Чек лист Заключение.docx"


def run(checklist_path, tpl, ctype, checked: dict, sources: dict = None):
    """{номер пункта: (статус, обоснование)} — номера, потому что у двух строк Акта
    («п.17» и подстрока «в соответствии с п.24») одинаковое название области."""
    items = checklists.parse_checklist(str(checklist_path(tpl)))
    all_texts = {**checked, **(D.SOURCES if sources is None else sources)}
    res = FC.run_checklist(ctype, items, checked, all_texts, doc_types.classify_all(all_texts))
    return {it["item_no"]: ("ok" if v["ok"] else "nok" if v["nok"] else "manual", v["reason"])
            for it, v in zip(items, res["checklist"])}


def akt(checklist_path, sources=None, **kw):
    return run(checklist_path, AKT_TPL, "akt", {"Акт.docx": D.akt(**kw)}, sources)


def zakl(checklist_path, akt_kw=None, sources=None, **kw):
    src = dict(D.SOURCES if sources is None else sources)
    src["Акт.docx"] = D.akt(**(akt_kw or {}))
    return run(checklist_path, ZAKL_TPL, "zakl", {"Заключение.docx": D.zakl(**kw)}, src)


def test_every_item_has_rule(checklist_path):
    for tpl, ctype in ((AKT_TPL, "akt"), (ZAKL_TPL, "zakl")):
        items = checklists.parse_checklist(str(checklist_path(tpl)))
        assert all(FC.rule_for(ctype, it) for it in items), tpl


def test_akt_all_ok(checklist_path):
    r = akt(checklist_path)
    manual = {n for n, (s, _) in r.items() if s == "manual"}
    assert manual == {13}, r  # подстрока «в соответствии с п.24» — на ручную проверку
    assert all(s == "ok" for n, (s, _) in r.items() if n != 13), r
    assert "в Разбивке детальнее: 28.13.14" in r[6][1]


def test_akt_type_and_area_problems(checklist_path):
    r = akt(checklist_path, types=("Первый инспекционный контроль",))
    assert r[2][0] == "nok" and "второй ИК" in r[2][1]
    r = akt(checklist_path, area="ремонту автомобилей и продаже запасных частей к ним в розницу")
    assert r[3][0] == "nok"


def test_akt_basis_sites_codes_team(checklist_path):
    src = dict(D.SOURCES)
    src["План аудита.docx"] = D.plan(team_expert="Кузнецов К.К.", period="28-30.07.2026",
                                     exclusions="п. 7.1.3.5 (абзац 1), п. 8.3")
    r = akt(checklist_path, sources=src, okved="28.13, 33.12, 71.12")
    assert r[6][0] == "nok" and "71.12" in r[6][1]           # коды
    assert r[7][0] == "nok"                                   # сроки
    assert r[9][0] == "nok" and "Кузнецов К.К." in r[9][1]    # состав ЭГ


def test_akt_compliance_and_counts(checklist_path):
    r = akt(checklist_path, sig="1", minor="0", nc_23_2="1")
    assert r[11][0] == "nok" and "частично" in r[11][1]       # п.14.4 при значительном НС
    r = akt(checklist_path, nc_23_2="3")
    assert r[15][0] == "nok"                                  # п.23.2 ≠ п.24


def test_akt_deadlines(checklist_path):
    # Акт2 = 30.07.2026 (чт), день подписания — первый: 5 р.д. — до 05.08, 20 р.д. — до 26.08
    r = akt(checklist_path, kd_deadline="27.08.2026", pkd_deadline="06.08.2026")
    assert r[12][0] == "nok" and "05.08.2026" in r[12][1]
    assert r[14][0] == "nok" and "26.08.2026" in r[14][1]


def test_akt_resertification_uses_cert_validity(checklist_path):
    types = ("Ресертификационный аудит",)
    src = {**D.SOURCES, "План аудита.docx": D.plan(types=("Ресертификационный аудит",))}
    # 40 р.д. с 30.07.2026 (день подписания — первый) — до 23.09.2026
    r = akt(checklist_path, sources=src, types=types, kd_deadline="23.09.2026", extend=False)
    assert r[14][0] == "manual" and "СЕРТ2" in r[14][1]       # сертификат не загружен
    r = akt(checklist_path, sources={**src, "Сертификат.pdf": D.SERT},
            types=types, kd_deadline="23.09.2026", extend=False)
    assert r[14][0] == "ok", r[14]


def test_akt_signatures_and_missing_sources(checklist_path):
    r = akt(checklist_path, signers=("И.И. Иванова",))
    assert r[17][0] == "nok" and "Петров П.П." in r[17][1]
    r = akt(checklist_path, sources={})
    assert r[1][0] == "nok" and "Договор" in r[1][1]
    assert r[2][0] == "nok" and "Приказ" in r[2][1]
    assert r[4][0] == "nok" and "План аудита" in r[4][1]


def test_zakl_all_ok(checklist_path):
    r = zakl(checklist_path)
    assert all(s == "ok" for s, _ in r.values()), r


def test_zakl_problems(checklist_path):
    r = zakl(checklist_path, types=(), approver="И.И. Иванова", approval="«25» августа 2026 г.",
             sign_date="25.08.2026", eliminated=False,
             decisions=("Выдать сертификат",))
    assert r[2][0] == "nok"                                    # утвердил сам РЭГ
    assert r[3][0] == "nok" and "19.08.2026" in r[3][1]        # позже Отчета КД + 5 р.д.
    assert r[4][0] == "nok" and "не отмечен вид аудита" in r[4][1]
    assert r[7][0] == "nok"                                    # «Несоответствия устранены»
    status, reason = r[9]                                      # п.4 по Акту: ИК + расширение
    assert status == "nok" and "«Подтвердить действие сертификата» не отмечено" in reason
    assert "вид аудита взят из Акта" in reason


def test_zakl_needs_otchet_kd(checklist_path):
    src = {k: v for k, v in D.SOURCES.items() if k != "Отчет по КД.pdf"}
    r = zakl(checklist_path, sources=src)
    assert r[3][0] == "nok" and "Отчет об устранении" in r[3][1]


def test_dki_package(samples_dir, checklist_path):
    """Пакет ДКИ из «ЧК ИИ/18.09». Договор ДКИ в образцах нет (лежит Договор УТЕКО-скан),
    Отчет по КД — скан без OCR. Находки на образцах: в п.17 Акта срок 06.08.2026, а 5 р.д.
    с даты подписания 30.07.2026 — до 05.08.2026; в п.1 Заключения не отмечен вид аудита."""
    import main

    def text(name):
        p = samples_dir / name
        return (main.extract_text_from_docx(str(p)) if p.suffix.lower() in (".docx", ".docm")
                else main.extract_text_from_pdf(str(p)))

    names = ["ИИ План аудита.docm", "ИИ Приказ об ЭГ.pdf",
             "ИИ Разбивка области сертификации по кодам ОКВЭД ДКИ (1).docx"]
    sources = {n: text(n) for n in names}
    akt_text = text("ИИ Акт инспекционный контроль.docx")

    r = run(checklist_path, AKT_TPL, "akt", {"Акт.docx": akt_text}, sources)
    assert r[1][0] == "nok"                    # Договор не загружен
    assert r[13][0] == "manual"                # подстрока п.24
    assert r[12][0] == "nok" and "05.08.2026" in r[12][1]
    assert r[14][0] == "ok", r[14]             # 26.08.2026 = 20 р.д. с даты подписания
    assert all(r[n][0] == "ok" for n in r if n not in (1, 12, 13)), r

    r = run(checklist_path, ZAKL_TPL, "zakl", {"Закл.docx": text("ИИ Закл.docx")},
            {**sources, "Акт.docx": akt_text, "ИИ Отчет по КД 2 этап.pdf": text("ИИ Отчет по КД 2 этап.pdf")})
    assert r[4][0] == "nok"                    # вид аудита в п.1 не отмечен
    assert r[3][0] in ("manual", "ok")         # Отчет по КД — скан; без OCR дата не читается
    assert all(r[n][0] == "ok" for n in r if n not in (3, 4)), r
