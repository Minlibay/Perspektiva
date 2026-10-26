import pytest

import checklists
import doc_types
import formula_checks as FC
from stage1_docs import AKT, OTCHET_KD, PKD, PLAN, loo_exp, loo_org, pn

TEMPLATES = {
    "pkd": "Чек лист ПКД.docx",
    "otchet_kd": "Чек лист Отчет КД.docx",
    "pn": "Чек лист ПН.docx",
    "loo_org": "Чек лист Лист оценки организацией.docx",
    "loo_exp": "Чек лист Лист оценки экспертами.docx",
}


def run(checklist_path, ctype, checked: dict, sources: dict = None):
    """Проверить чек-лист ctype; возвращает {название пункта: (статус, обоснование)}."""
    items = checklists.parse_checklist(str(checklist_path(TEMPLATES[ctype])))
    all_texts = {**checked, **(sources if sources is not None else {"Акт.docx": AKT, "План.docx": PLAN})}
    res = FC.run_checklist(ctype, items, checked, all_texts, doc_types.classify_all(all_texts))
    out = {}
    for it, v in zip(items, res["checklist"]):
        status = "ok" if v["ok"] else "nok" if v["nok"] else "manual"
        out[" ".join(it["area"].split())] = (status, v["reason"])
    return out


def statuses(result: dict) -> list:
    return [s for s, _ in result.values()]


@pytest.mark.parametrize("ctype", list(TEMPLATES))
def test_every_item_has_rule(checklist_path, ctype):
    items = checklists.parse_checklist(str(checklist_path(TEMPLATES[ctype])))
    assert all(FC.rule_for(ctype, it) for it in items)


def test_pkd_ok(checklist_path):
    r = run(checklist_path, "pkd", {"ПКД.docx": PKD})
    assert statuses(r) == ["ok", "ok", "ok"], r


def test_pkd_late_and_rows_mismatch(checklist_path):
    late = PKD.replace("«03» августа 2026 г.", "«07» августа 2026 г.").replace(
        "2 | Нет записи об обучении | М | 18.08.2026г. | -\n", "")
    r = run(checklist_path, "pkd", {"ПКД.docx": late})
    assert r["Дата утверждения"][0] == "nok" and "06.08.2026" in r["Дата утверждения"][1]
    assert r["Количество строк в таблице"][0] == "nok"


def test_otchet_kd_requires_full_name(checklist_path):
    r = run(checklist_path, "otchet_kd", {"Отчет по КД.pdf": OTCHET_KD})
    assert statuses(r) == ["ok", "ok", "ok"], r
    short = OTCHET_KD.replace("Общество с ограниченной ответственностью «Ромашка»\nнаименование",
                              "ООО «Ромашка»\nнаименование")
    r = run(checklist_path, "otchet_kd", {"Отчет по КД.pdf": short})
    assert r["Название организации полностью"][0] == "nok"
    assert "сокращённое" in r["Название организации полностью"][1]


def test_unreadable_scan_goes_to_manual_check(checklist_path):
    r = run(checklist_path, "otchet_kd", {"Отчет по КД.pdf": "КД1КД2"})
    assert statuses(r) == ["manual"] * 3
    assert all(reason.startswith("ПРОВЕРИТЬ ВРУЧНУЮ") for _, reason in r.values())


def test_missing_akt_is_nok(checklist_path):
    r = run(checklist_path, "pkd", {"ПКД.docx": PKD}, sources={})
    assert statuses(r) == ["nok"] * 3
    assert "Не загружен Акт" in r["Название организации"][1]


def test_pn_ok_two_protocols(checklist_path):
    r = run(checklist_path, "pn", {"ПН.docx": pn(n=2)})
    assert statuses(r) == ["ok"] * 6, r


def test_pn_problems(checklist_path):
    r = run(checklist_path, "pn", {"ПН.docx": pn(date="31.07.2026", leader="П.П. Петров", sto=False)})
    assert r["Дата"][0] == "nok"
    assert r["ФИО Руководителя ЭГ"][0] == "nok" and "Петров П.П." in r["ФИО Руководителя ЭГ"][1]
    assert r["Пункты нормативного документа"][0] == "nok"
    assert r["Количество протоколов"][0] == "nok"  # 1 протокол, в Акте 2 несоответствия


def test_pn_org_falls_back_to_akt_without_plan(checklist_path):
    r = run(checklist_path, "pn", {"ПН.docx": pn(n=2)}, sources={"Акт.docx": AKT})
    assert r["Название организации"][0] == "ok"
    assert "План аудита не загружен" in r["Название организации"][1]


def test_loo_org_ok(checklist_path):
    r = run(checklist_path, "loo_org", {"Лист.docx": loo_org()})
    assert statuses(r) == ["ok"] * 6, r


def test_loo_org_problems(checklist_path):
    r = run(checklist_path, "loo_org", {"Лист.docx": loo_org(
        role="Член экспертной группы", marks_per_row=0, date="31.07.2026", signer="Петров П.П.")})
    assert r["Оцениваемый Эксперт"][0] == "nok" and "в Акте — руководитель ЭГ" in r["Оцениваемый Эксперт"][1]
    assert r["Отметки в таблице"][0] == "nok" and "13 из 13" in r["Отметки в таблице"][1]
    assert r["Дата"][0] == "nok"
    assert r["ФИО Должность"][0] == "nok"


def test_loo_org_two_marks_in_row(checklist_path):
    r = run(checklist_path, "loo_org", {"Лист.docx": loo_org(marks_per_row=2)})
    assert r["Отметки в таблице"][0] == "nok" and "несколькими" in r["Отметки в таблице"][1]


def test_loo_exp(checklist_path):
    r = run(checklist_path, "loo_exp", {"Лист.docx": loo_exp()})
    assert statuses(r) == ["ok"] * 6, r
    r = run(checklist_path, "loo_exp", {"Лист.docx": loo_exp(standard="ГОСТ Р ИСО 9001")})
    assert r["Отметка в окошке"][0] == "nok"
    # кандидат, отмеченный как эксперт
    r = run(checklist_path, "loo_exp", {"Лист.docx": loo_exp(person="Кузнецов К.К.", role="Эксперт")})
    assert r["Оцениваемый Эксперт"][0] == "nok" and "кандидат в эксперты" in r["Оцениваемый Эксперт"][1]


def test_several_sheets_aggregated(checklist_path):
    r = run(checklist_path, "loo_org", {"Лист 1.docx": loo_org(),
                                        "Лист 2.docx": loo_org(person="Петров П.П.",
                                                               role="Член экспертной группы",
                                                               marks_per_row=0)})
    status, reason = r["Отметки в таблице"]
    assert status == "nok" and "Лист 1.docx:" in reason and "Лист 2.docx:" in reason
    assert r["Оцениваемый Эксперт"][0] == "ok"


def test_dki_package(samples_dir, checklist_path):
    """Пакет ДКИ из «ЧК ИИ/18.09». Листы оценки в образцах не заполнены (таблицы
    оценок пустые) — это и должна показать проверка."""
    import main

    def text(name):
        p = samples_dir / name
        return (main.extract_text_from_docx(str(p)) if p.suffix.lower() in (".docx", ".docm")
                else main.extract_text_from_pdf(str(p)))

    sources = {n: text(n) for n in ("ИИ Акт инспекционный контроль.docx", "ИИ План аудита.docm")}
    assert statuses(run(checklist_path, "pkd", {"ИИ План КД.docx": text("ИИ План КД.docx")}, sources)) == ["ok"] * 3
    assert statuses(run(checklist_path, "pn", {"ПН.docx": text("ИИ Протоколы несоотв.docx")}, sources)) == ["ok"] * 6
    r = run(checklist_path, "loo_org", {"ЛОО.docx": text("ИИ Лист Оценки  организацией.docx")}, sources)
    assert statuses(r) == ["ok"] * 5 + ["nok"]
    r = run(checklist_path, "loo_exp", {"ЛОЭ.docx": text("ИИ Лист оценки экспертами.docx")}, sources)
    assert statuses(r) == ["ok"] * 5 + ["nok"]
