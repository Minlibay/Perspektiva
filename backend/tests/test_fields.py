from datetime import date

import fields as F
from stage1_docs import AKT, OTCHET_KD, PKD, PLAN, loo_exp, loo_org, pn


def test_dates_and_periods():
    assert F.find_dates("от 30.07.2026 и «12» августа 2026 г.") == [date(2026, 7, 30), date(2026, 8, 12)]
    assert F.find_period("6 Сроки | 27-30.07.2026") == (date(2026, 7, 27), date(2026, 7, 30))
    assert F.find_period("с 27.07.2026 по 30.07.2026") == (date(2026, 7, 27), date(2026, 7, 30))
    assert F.find_period("30.07.2026") == (date(2026, 7, 30), date(2026, 7, 30))
    assert F.add_workdays(date(2026, 2, 10), -5) == date(2026, 2, 3)
    # сроки Акта: день подписания — первый рабочий день
    assert F.workday_deadline(date(2026, 7, 30), 5) == date(2026, 8, 5)
    assert F.workday_deadline(date(2026, 7, 30), 20) == date(2026, 8, 26)
    assert F.workday_deadline(date(2026, 8, 1), 5) == date(2026, 8, 7)  # суббота — с понедельника


def test_org_full_and_short_forms_are_equal():
    full = F.find_org("Общество с ограниченной ответственностью «Ромашка»")
    short = F.find_org("(ООО «Ромашка»)")
    assert full["key"] == short["key"] == ("ооо", "ромашка")
    assert full["full_form"] and not short["full_form"]
    # орган по сертификации пропускается, вложенные кавычки не ломают разбор
    assert F.find_org("ОС СМК ООО «КЦ «Перспектива»» и АО «Альфа»")["key"] == ("ао", "альфа")
    assert F.find_org("УТВЕРЖДАЮ ООО «Ромашка» … Общество с ограниченной ответственностью «Ромашка»",
                      last=True)["full_form"]
    # Сводный акт ПТС: у «КЦ «Перспектива»» нет внешней закрывающей кавычки — раньше
    # название растягивалось на страницу текста до следующей «»»
    unbalanced = "ОС СМК ООО «КЦ «Перспектива» 1. Область … АКЦИОНЕРНОЕ ОБЩЕСТВО «ПТС» (АО «ПТС»)"
    assert F.find_org(unbalanced)["key"] == ("ао", "птс")
    assert F.find_org("АО «Завод «Прогресс» им. Ленина»")["key"] == ("ао", "заводпрогрессимленина")


def test_persons():
    assert F.find_persons("Сукорская О.А. Гревцов К.В.") == ["сукорская оа", "гревцов кв"]
    assert F.find_persons("| О.А. Сукорская Акт3 |  | А.Б.Татищев") == ["сукорская оа", "татищев аб"]
    assert F.find_persons("Технический управляющий Татищев А.Б.") == ["татищев аб"]
    assert F.person_label("сукорская оа") == "Сукорская О.А."


def test_akt_fields():
    f = F.akt_fields(AKT)
    assert f["org"]["key"] == ("ооо", "ромашка") and f["org"]["full_form"]
    assert f["period"] == (date(2026, 7, 27), date(2026, 7, 30))
    assert f["team"]["leader"] == ["иванова ии"]
    assert f["team"]["experts"] == ["петров пп"]
    assert f["team"]["candidates"] == ["кузнецов кк"]
    assert f["representative"] == "сидоров сс"
    assert f["date"] == date(2026, 7, 30) and f["leader_sign"] == "иванова ии"
    assert (f["nc_total"], f["nc_significant"], f["nc_minor"]) == (2, 0, 2)
    assert f["pkd_deadline"] == date(2026, 8, 6)
    assert f["kd_deadline"] == date(2026, 8, 26)


def test_akt_nc_total_falls_back_to_results_table():
    f = F.akt_fields(AKT.replace("В ходе настоящего аудита выявлено 2 несоответствий.", ""))
    assert f["nc_total"] == 2


def test_plan_pkd_kd_fields():
    assert F.plan_fields(PLAN)["org"]["key"] == ("ооо", "ромашка")
    pkd = F.approved_doc_fields(PKD)
    assert pkd["org"]["full_form"] and pkd["date"] == date(2026, 8, 3) and pkd["rows"] == 2
    kd = F.approved_doc_fields(OTCHET_KD)
    # ближайшее к подписи полное наименование, а не «ООО» из блока «УТВЕРЖДАЮ»
    assert kd["org"]["full_form"] and kd["date"] == date(2026, 8, 12) and kd["rows"] == 2
    assert F.approved_doc_fields("скан без текста")["rows"] is None


def test_pn_fields():
    f = F.pn_fields(pn(n=2))
    assert len(f["protocols"]) == 2
    p = f["protocols"][0]
    assert p["org"]["key"] == ("ооо", "ромашка") and p["date"] == date(2026, 7, 30)
    assert p["leader"] == "иванова ии" and p["has_sto"]
    assert f["version"] == "вер. 1 от 13.01.25"


def test_loo_fields():
    f = F.loo_fields(loo_org())
    assert f["org"]["key"] == ("ооо", "ромашка")
    assert f["period"] == (date(2026, 7, 27), date(2026, 7, 30))
    assert f["person"] == "иванова ии" and f["roles"] == ["leader"]
    assert f["date"] == date(2026, 7, 30) and f["signer"] == "сидоров сс"
    assert len(f["rating_rows"]) == 13 and set(f["rating_rows"].values()) == {1}

    e = F.loo_fields(loo_exp())
    assert e["standards"] == ["СТО Газпром 9001"] and e["roles"] == ["experts"]
    assert e["person"] == "петров пп" and len(e["rating_rows"]) == 7
