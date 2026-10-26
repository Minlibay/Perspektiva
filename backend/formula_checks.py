"""
Чек-листы, проверяемые «по формулам» (docs/ЧЕК-ЛИСТЫ.md §4.3–4.9): этап 1 — ПКД, Отчет КД,
Протоколы НС, Листы оценки; этап 2 — Акт и Заключение; этап 3 — Отчет 1 этапа и Заключение
1 этапа.

Формулы шаблонов — сравнение полей документов («ПКД1=А20 Акта», «КД2≤Акт6»), поэтому
вердикт выносит код по полям из fields.py, GigaChat не вызывается. Правила:
- поле не распознано (плохой OCR, нестандартная форма) — «ПРОВЕРИТЬ ВРУЧНУЮ», ОК/NOK
  пустые: нераспознанное поле — не нарушение;
- не загружен документ-источник (Акт) — NOK, как в пайплайне Плана.
Коды без разметки в образцах трактуются по перекрёстным ссылкам в самих образцах
(ASSUMPTIONS ниже, docs/ЧЕК-ЛИСТЫ.md §5).
"""

import re
from datetime import date
from typing import Callable, Optional

import doc_types
import fields as F

# Трактовки кодов, которых нет в разметке образцов «ЧК ИИ/18.09». Основание — перекрёстные
# ссылки в образцах: в Протоколе НС у даты стоит «Акт 2», у ФИО РЭГ — «Акт3»; в Отчете
# по КД полное наименование помечено «КД1».
ASSUMPTIONS = {
    "pn": "И2, И3, И5 Акта — это Акт2 (дата Акта), Акт3 (ФИО РЭГ), Акт5 (число несоответствий).",
    "pkd": "«ПКД2 = Акт-ПКД1» проверяется как «утверждён не позднее срока из п.17 Акта»; "
           "И5 Акта — Акт5.",
    "otchet_kd": "«1Отчет2 = А20» — это КД1 (наименование организации в Отчете по КД).",
    "loo": "«А1 Акта» — наименование организации из Акта (Акт1/А20); «ЛОО332» — ЛОО33.",
    "akt": "«= Плану 2 этапа» — сверка с загруженным Планом аудита; СЕРТ2 — дата окончания срока действия сертификата («по …»); рабочие дни п.17/п.18 считаются с даты подписания Акта — день подписания первый (решение методолога 12.10.2026); подстрока «в соответствии с п.24» — на ручную проверку с данными п.14.4/п.17/п.24.",
    "zakl": "п.4 проверяется по подписям отметок, а не по кодам З10/З13/З14 (в шаблоне и образце коды не совпадают); «З2=Говорков → З1≠Говорков» обобщено до «утверждающий ≠ руководитель ЭГ».",
    "otchet1": "«1О7 ≤ 1О8 + 5 р.д.» — от последнего дня аудита, рабочие дни со следующего дня; реквизиты Основания (1О6) сверяются с Договором и всеми загруженными доп. соглашениями.",
    "zakl1": "Образца Заключения 1 этапа нет — разбор по аналогии с Заключением того же ОС; правило «1Закл3≠Говорков → 1Закл2=Говорков, иначе Сукорская» обобщено до «утверждающий ≠ руководитель ЭГ».",
}


def _v(status: str, reason: str) -> dict:
    return {"status": status, "reason": reason}


def _ok(reason):
    return _v("ok", reason)


def _nok(reason):
    return _v("nok", reason)


def _manual(reason):
    return _v("manual", reason)


def _combine(results: list) -> dict:
    """Итог по нескольким документам / протоколам: NOK, если хоть один NOK; иначе
    ручная проверка, если хоть один не распознан; иначе OK. results — [(метка, вердикт)]."""
    if not results:
        return _manual("проверяемый документ не найден")
    if len(results) == 1:
        return results[0][1]
    status = ("nok" if any(v["status"] == "nok" for _, v in results)
              else "manual" if any(v["status"] == "manual" for _, v in results) else "ok")
    return _v(status, "; ".join(f"{label}: {v['reason']}" for label, v in results))


class Ctx:
    """Документы для проверки: проверяемые (checked — {файл: текст}) и источники."""

    def __init__(self, checked: dict, all_texts: dict, doc_type_map: dict, formula: str = ""):
        self.checked = checked
        self.all_texts = all_texts
        self.doc_type_map = doc_type_map
        self.formula = formula
        self._cache = {}

    def source(self, dtypes, parse: Callable, pick: Optional[Callable] = None):
        """(файл, поля) документа-источника одного из типов dtypes (строка или кортеж,
        в порядке предпочтения) или (None, None). pick(поля) — ключ выбора среди
        нескольких документов (берётся максимальный); по умолчанию — первый."""
        dtypes = (dtypes,) if isinstance(dtypes, str) else tuple(dtypes)
        if dtypes not in self._cache:
            found = (None, None)
            for dtype in dtypes:
                parsed = [(n, parse(self.all_texts[n])) for n in doc_types.files_of_type(self.doc_type_map, dtype)]
                if parsed:
                    found = max(parsed, key=lambda x: pick(x[1])) if pick else parsed[0]
                    break
            self._cache[dtypes] = found
        return self._cache[dtypes]

    @property
    def akt(self):
        # Актов может быть несколько (текущий и предыдущего аудита) — берём самый поздний
        return self.source("akt", F.akt_fields, pick=lambda f: f["date"] or date.min)

    @property
    def plan(self):
        return self.source("plan", F.plan_fields)

    def each(self, parse: Callable, check: Callable) -> dict:
        """check(поля документа) для каждого проверяемого документа."""
        return _combine([(name, check(parse(text))) for name, text in self.checked.items()])


NO_AKT = "Не загружен Акт по результатам аудита — сверять не с чем."


def _org(doc_org, ref_org, src: str, require_full: bool = False) -> dict:
    """Сверка наименования организации. src — источник в именительном падеже: «Акт (А20)»."""
    if not doc_org:
        return _manual("наименование организации в документе не распознано")
    if not ref_org:
        return _manual(f"не распознано наименование организации для сверки — {src}")
    if doc_org["key"] != ref_org["key"]:
        return _nok(f"«{doc_org['name']}» не совпадает: {src} — «{ref_org['name']}»")
    if require_full and not doc_org["full_form"]:
        return _nok(f"указано сокращённое наименование «{doc_org['name']}», требуется полное "
                    f"({src} — «{ref_org['name']}»)")
    return _ok(f"«{doc_org['name']}» совпадает ({src})")


def _same_date(doc_date, ref_date, what: str, src: str) -> dict:
    if not doc_date:
        return _manual(f"{what} в документе не распознана")
    if not ref_date:
        return _manual(f"не распознана дата для сверки — {src}")
    if doc_date != ref_date:
        return _nok(f"{what} {F.fmt_date(doc_date)} ≠ {src}: {F.fmt_date(ref_date)}")
    return _ok(f"{what} {F.fmt_date(doc_date)} совпадает ({src})")


def _not_later(doc_date, deadline, what: str, deadline_label: str) -> dict:
    if not doc_date:
        return _manual(f"{what} в документе не распознана")
    if not deadline:
        return _manual(f"не распознан {deadline_label}")
    if doc_date > deadline:
        return _nok(f"{what} {F.fmt_date(doc_date)} позже срока: {deadline_label} — {F.fmt_date(deadline)}")
    return _ok(f"{what} {F.fmt_date(doc_date)} не позднее срока ({deadline_label} — {F.fmt_date(deadline)})")


def _rows(rows, nc_total) -> dict:
    if rows is None:
        return _manual("таблица несоответствий не распознана")
    if nc_total is None:
        return _manual("число несоответствий в Акте (п.23.2 / п.24) не распознано")
    if rows != nc_total:
        return _nok(f"строк в таблице {rows}, а в Акте выявлено несоответствий: {nc_total}")
    return _ok(f"строк в таблице {rows} — столько же несоответствий в Акте")


# ---------- ПКД и Отчет по КД ----------

def _approved_org(ctx: Ctx, require_full: bool) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return ctx.each(F.approved_doc_fields, lambda f: _org(f["org"], akt["org"], "Акт (А20)", require_full))


def _pkd_date(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return ctx.each(F.approved_doc_fields, lambda f: _not_later(
        f["date"], akt["pkd_deadline"], "Дата утверждения ПКД", "срок представления плана КД (п.17 Акта)"))


def _kd_date(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return ctx.each(F.approved_doc_fields, lambda f: _not_later(
        f["date"], akt["kd_deadline"], "Дата утверждения отчета", "срок представления отчета (п.18 Акта, Акт6)"))


def _approved_rows(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return ctx.each(F.approved_doc_fields, lambda f: _rows(f["rows"], akt["nc_total"]))


# ---------- Протоколы регистрации несоответствий ----------

def _protocols(ctx: Ctx) -> list:
    """[(метка, поля протокола)] по всем протоколам всех проверяемых файлов."""
    out = []
    for name, text in ctx.checked.items():
        prots = F.pn_fields(text)["protocols"]
        for i, p in enumerate(prots, 1):
            out.append((name if len(prots) == 1 else f"{name}, протокол {i}", p))
    return out


def _pn_each(ctx: Ctx, check: Callable) -> dict:
    prots = _protocols(ctx)
    if not prots:
        return _manual("протоколы регистрации несоответствий в документе не распознаны")
    return _combine([(label, check(p)) for label, p in prots])


def _pn_org(ctx: Ctx) -> dict:
    plan_file, plan = ctx.plan
    if plan_file:
        ref, label = plan["org"], "План аудита (П1)"
    else:
        akt_file, akt = ctx.akt
        if not akt_file:
            return _nok("Не загружен ни План аудита, ни Акт — сверять наименование не с чем.")
        ref, label = akt["org"], "Акт; План аудита не загружен"
    return _pn_each(ctx, lambda p: _org(p["org"], ref, label))


def _pn_date(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return _pn_each(ctx, lambda p: _same_date(p["date"], akt["date"], "Дата протокола", "дата Акта (Акт2)"))


def _same_person(doc_person, ref_person, what: str, src: str) -> dict:
    if not doc_person:
        return _manual(f"{what} в документе не распознано")
    if not ref_person:
        return _manual(f"не распознано ФИО для сверки — {src}")
    if doc_person != ref_person:
        return _nok(f"{what}: {F.person_label(doc_person)} ≠ {src}: {F.person_label(ref_person)}")
    return _ok(f"{what}: {F.person_label(doc_person)} — совпадает ({src})")


def _akt_leader(akt: dict) -> Optional[str]:
    return akt["leader_sign"] or next(iter(akt["team"].get("leader") or []), None)


def _pn_leader(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return _pn_each(ctx, lambda p: _same_person(p["leader"], _akt_leader(akt), "ФИО руководителя ЭГ", "Акт (Акт3)"))


def _pn_sto(ctx: Ctx) -> dict:
    return _pn_each(ctx, lambda p: _ok("указан СТО Газпром 9001-2018") if p["has_sto"]
                    else _nok("в пунктах нормативного документа нет СТО Газпром 9001-2018"))


def _pn_count(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    n = len(_protocols(ctx))
    if not n:
        return _manual("протоколы регистрации несоответствий в документе не распознаны")
    if akt["nc_total"] is None:
        return _manual("число несоответствий в Акте (п.23.2 / п.24) не распознано")
    if n != akt["nc_total"]:
        return _nok(f"протоколов {n}, а в Акте выявлено несоответствий: {akt['nc_total']}")
    return _ok(f"протоколов {n} — столько же несоответствий в Акте")


def _norm_version(s: str) -> str:
    return re.sub(r"[\sг.]+$", "", re.sub(r"\s+", "", (s or "").lower()))


def _pn_version(ctx: Ctx) -> dict:
    m = re.search(r"вер\.\s*\d+\s*от\s*[\d.]+", ctx.formula)
    if not m:
        return _manual("в шаблоне не указана ожидаемая версия формы")
    expected = m.group(0)

    def check(f):
        if not f["version"]:
            return _manual("отметка о версии формы не найдена")
        if _norm_version(f["version"]) != _norm_version(expected):
            return _nok(f"версия формы «{f['version']}», требуется «{expected}»")
        return _ok(f"версия формы «{f['version']}»")

    return ctx.each(F.pn_fields, check)


# ---------- Листы оценки ----------

def _loo_org(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return ctx.each(F.loo_fields, lambda f: _org(f["org"], akt["org"], "Акт"))


def _loo_period(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)

    def check(f):
        if not f["period"]:
            return _manual("дата аудита в листе не распознана")
        if not akt["period"]:
            return _manual("сроки проведения аудита в Акте (А8) не распознаны")
        if f["period"] != akt["period"]:
            return _nok(f"дата аудита {F.fmt_period(f['period'])} ≠ Акт (А8): {F.fmt_period(akt['period'])}")
        return _ok(f"дата аудита {F.fmt_period(f['period'])} совпадает с Актом (А8)")

    return ctx.each(F.loo_fields, check)


_ROLE_SOURCES = {  # роль в листе → роли в п.8 Акта
    "leader": ("leader",),
    "experts": ("experts",),
    "candidates": ("candidates", "tech"),
}


def _loo_person(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    team = akt["team"]

    def check(f):
        if not f["person"]:
            return _manual("оцениваемый эксперт в листе не распознан")
        if not f["roles"]:
            return _nok(f"не отмечена роль оцениваемого ({F.person_label(f['person'])})")
        if len(f["roles"]) > 1:
            return _nok("отмечено несколько ролей: " + ", ".join(F.ROLE_TITLES[r] for r in f["roles"]))
        role = f["roles"][0]
        in_role = [p for r in _ROLE_SOURCES[role] for p in team.get(r, [])]
        if f["person"] in in_role:
            return _ok(f"{F.person_label(f['person'])} — {F.ROLE_TITLES[role]}, как в составе ЭГ Акта")
        actual = next((F.ROLE_TITLES[r] for r, ps in team.items() if f["person"] in ps), None)
        if actual:
            return _nok(f"{F.person_label(f['person'])} отмечен(а) как {F.ROLE_TITLES[role]}, "
                        f"а в Акте — {actual}")
        if not any(team.values()):
            return _manual("состав ЭГ в Акте (А21) не распознан")
        return _nok(f"{F.person_label(f['person'])} нет в составе ЭГ Акта")

    return ctx.each(F.loo_fields, check)


def _loo_date(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return ctx.each(F.loo_fields, lambda f: _same_date(f["date"], akt["date"], "Дата листа", "дата Акта (Акт2)"))


def _loo_signer(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return ctx.each(F.loo_fields, lambda f: _same_person(
        f["signer"], akt["representative"], "ФИО подписавшего", "уполномоченный представитель в Акте (А23)"))


def _loo_marks(ctx: Ctx) -> dict:
    def check(f):
        rows = f["rating_rows"]
        if not rows:
            return _manual("таблица оценок в листе не распознана")
        empty = [k for k, n in rows.items() if n == 0]
        multi = [k for k, n in rows.items() if n > 1]
        if not empty and not multi:
            return _ok(f"во всех {len(rows)} строках оценок по одной отметке")
        parts = []
        if empty:
            parts.append(f"без отметки строк: {len(empty)} из {len(rows)} (например, «{empty[0][:60]}»)")
        if multi:
            parts.append(f"с несколькими отметками: {len(multi)} (например, «{multi[0][:60]}»)")
        return _nok("; ".join(parts))

    return ctx.each(F.loo_fields, check)


def _loo_standard(ctx: Ctx) -> dict:
    def check(f):
        std = f["standards"]
        if not std:
            return _nok("не отмечен стандарт (требуется «СТО Газпром 9001»)")
        if len(std) > 1:
            return _nok("отмечено несколько стандартов: " + ", ".join(std))
        if "сто газпром 9001" not in std[0].lower():
            return _nok(f"отмечен «{std[0]}», требуется «СТО Газпром 9001»")
        return _ok("отмечен «СТО Газпром 9001»")

    return ctx.each(F.loo_fields, check)


# ---------- Акт по результатам аудита (проверяемый документ — Акт) ----------

def _types_label(types: set) -> str:
    return ", ".join(sorted(types)) or "ничего не отмечено"


def _need(ctx: Ctx, dtypes, parse: Callable, missing: str):
    """Источник или вердикт NOK «не загружен …»: (файл, поля, None) | (None, None, вердикт)."""
    name, f = ctx.source(dtypes, parse)
    return (name, f, None) if name else (None, None, _nok(f"Не загружен {missing} — сверять не с чем."))


def _akt_org(ctx: Ctx) -> dict:
    _, contract, err = _need(ctx, ("dogovor", "dop_soglashenie"), F.contract_fields, "Договор")
    if err:
        return err
    return ctx.each(F.akt_fields, lambda f: _org(f["org"], contract["org"], "Договор (Д2)", require_full=True))


def _akt_type(ctx: Ctx) -> dict:
    _, prikaz, err = _need(ctx, "prikaz", F.prikaz_fields, "Приказ о назначении ЭГ")
    if err:
        return err
    _, plan = ctx.plan

    def check(f):
        if not f["types"]:
            return _nok("в п.1 Акта не отмечен вид аудита")
        main = f["types"] - {F.EXTENSION}
        if not prikaz["audit_type"]:
            return _manual(f"в Акте отмечено: {_types_label(f['types'])}; вид аудита в Приказе не распознан")
        if main != {prikaz["audit_type"]}:
            return _nok(f"в Акте отмечено: {_types_label(f['types'])}, в Приказе — {prikaz['audit_type']}")
        reason = f"вид аудита «{prikaz['audit_type']}» совпадает с Приказом"
        if plan is None:
            return _ok(reason + "; План аудита не загружен — с ним не сверено")
        if plan["types"] and plan["types"] - {F.EXTENSION} != main:
            return _nok(f"в Акте и Приказе — {prikaz['audit_type']}, в Плане отмечено: {_types_label(plan['types'])}")
        if plan["types"] and F.EXTENSION in f["types"] and F.EXTENSION not in plan["types"]:
            return _nok("в Акте отмечено расширение области, в Плане — нет")
        return _ok(reason + (" и Планом" if plan["types"] else "; в Плане отметки вида аудита не распознаны"))

    return ctx.each(F.akt_fields, check)


def _akt_area(ctx: Ctx) -> dict:
    _, razb, err = _need(ctx, "razbivka", F.razbivka_fields, "файл «Разбивка области по кодам ОКВЭД»")
    if err:
        return err
    _, plan = ctx.plan
    refs = [("Разбивка (ОБ1)", razb["area"])] + ([("План (П4)", plan["area"])] if plan else [])

    def check(f):
        words = F.word_set(f["area"])
        if not words:
            return _manual("область применения СМК в Акте (А30) не распознана")
        parts, worst = [], 1.0
        for label, text in refs:
            if not F.word_set(text):
                parts.append(f"{label} — область не распознана")
                worst = min(worst, 0.6)  # нечем сверить — на ручную проверку
                continue
            jac = F.jaccard(words, F.word_set(text))
            worst = min(worst, jac)
            parts.append(f"{label} — совпадение слов {int(jac * 100)}%")
        reason = "формулировка области: " + "; ".join(parts)
        if worst >= 0.7:
            return _ok(reason)
        if worst < 0.5:
            return _nok(reason)
        return _manual(reason + " — сверьте формулировку")

    return ctx.each(F.akt_fields, check)


def _plan_or_nok(ctx: Ctx):
    name, plan = ctx.plan
    return (plan, None) if name else (None, _nok("Не загружен План аудита — сверять не с чем."))


def _akt_basis(ctx: Ctx) -> dict:
    plan, err = _plan_or_nok(ctx)
    if err:
        return err

    def check(f):
        nums, dates = f["basis"]
        p_nums, p_dates = plan["basis"]
        if not nums and not dates:
            return _manual("реквизиты основания в Акте (А31) не распознаны")
        if not p_nums and not p_dates:
            return _manual("реквизиты основания в Плане (П2) не распознаны")
        diff = []
        if nums != p_nums:
            diff.append(f"номера: Акт {sorted(nums)}, План {sorted(p_nums)}")
        if dates != p_dates:
            diff.append(f"даты: Акт {sorted(map(F.fmt_date, dates))}, План {sorted(map(F.fmt_date, p_dates))}")
        if diff:
            return _nok("основание расходится с Планом (П2) — " + "; ".join(diff))
        return _ok(f"номера ({', '.join(sorted(nums))}) и даты основания совпадают с Планом (П2)")

    return ctx.each(F.akt_fields, check)


def _akt_sites(ctx: Ctx) -> dict:
    plan, err = _plan_or_nok(ctx)
    if err:
        return err

    def check(f):
        postals, words = f["sites"]
        p_postals, p_words = plan["sites"]
        if not words:
            return _manual("адреса площадок в Акте (А32) не распознаны")
        if not p_words:
            return _manual("адреса площадок в Плане (П3) не распознаны")
        if postals != p_postals:
            return _nok(f"индексы площадок: Акт {sorted(postals) or '—'}, План {sorted(p_postals) or '—'}")
        jac = F.jaccard(words, p_words)
        if jac >= 0.8:
            return _ok(f"адреса площадок совпадают с Планом (П3), индексы: {', '.join(sorted(postals)) or '—'}")
        return _manual(f"адреса площадок совпадают с Планом на {int(jac * 100)}% — сверьте "
                       f"(только в Акте: {', '.join(sorted(words - p_words)[:5]) or '—'}; "
                       f"только в Плане: {', '.join(sorted(p_words - words)[:5]) or '—'})")

    return ctx.each(F.akt_fields, check)


def _akt_okved(ctx: Ctx) -> dict:
    _, razb, err = _need(ctx, "razbivka", F.razbivka_fields, "файл «Разбивка области по кодам ОКВЭД»")
    if err:
        return err

    def check(f):
        if not f["okved"]:
            return _manual("коды ОКВЭД в Акте (А33) не распознаны")
        if not razb["okved"]:
            return _manual("коды ОКВЭД в Разбивке (ОБ2) не распознаны")
        akt = {F.okved_group(c) for c in f["okved"]}
        ref = {F.okved_group(c) for c in razb["okved"]}
        if akt != ref:
            return _nok(f"коды расходятся с Разбивкой (ОБ2): только в Акте — {sorted(akt - ref) or '—'}, "
                        f"только в Разбивке — {sorted(ref - akt) or '—'}")
        detail = [c for c in razb["okved"] if c not in f["okved"]]
        return _ok(f"коды ОКВЭД ({len(akt)}) совпадают с Разбивкой (ОБ2) на уровне группировок"
                   + (f"; в Разбивке детальнее: {', '.join(detail)}" if detail else ""))

    return ctx.each(F.akt_fields, check)


def _akt_period(ctx: Ctx) -> dict:
    plan, err = _plan_or_nok(ctx)
    if err:
        return err

    def check(f):
        if not f["period"] or not plan["period"]:
            return _manual("сроки проведения аудита в " + ("Акте (А8)" if not f["period"] else "Плане (П4)")
                           + " не распознаны")
        if f["period"] != plan["period"]:
            return _nok(f"сроки в Акте {F.fmt_period(f['period'])} ≠ План: {F.fmt_period(plan['period'])}")
        return _ok(f"сроки {F.fmt_period(f['period'])} совпадают с Планом")

    return ctx.each(F.akt_fields, check)


def _akt_representative(ctx: Ctx) -> dict:
    plan, err = _plan_or_nok(ctx)
    if err:
        return err
    return ctx.each(F.akt_fields, lambda f: _same_person(
        f["representative"], plan["representative"], "Уполномоченный представитель", "План (П5)"))


def _team_label(team: dict) -> str:
    return "; ".join(f"{F.ROLE_TITLES[r]}: {', '.join(F.person_label(p) for p in ps)}"
                     for r, ps in team.items() if ps) or "—"


def _akt_team(ctx: Ctx) -> dict:
    plan, err = _plan_or_nok(ctx)
    if err:
        return err
    norm = lambda team: {r: sorted(ps) for r, ps in team.items() if ps}

    def check(f):
        a, p = norm(f["team"]), norm(plan["team"])
        if not a:
            return _manual("состав ЭГ в Акте (А21) не распознан")
        if not p:
            return _manual("состав ЭГ в Плане (П6) не распознан")
        if a != p:
            return _nok(f"состав ЭГ расходится — Акт: {_team_label(f['team'])}; План: {_team_label(plan['team'])}")
        return _ok(f"состав ЭГ совпадает с Планом: {_team_label(f['team'])}")

    return ctx.each(F.akt_fields, check)


def _akt_exclusions(ctx: Ctx) -> dict:
    plan, err = _plan_or_nok(ctx)
    if err:
        return err

    def check(f):
        if f["exclusions"] is None:
            return _manual("раздел «Исключения из требований стандартов» в Акте (А34) не найден")
        if plan["exclusions"] is None:
            return _manual("раздел «Исключения из требований стандартов» в Плане (П7) не найден")
        a, p = set(f["exclusions"]), set(plan["exclusions"])
        if a != p:
            return _nok(f"исключения расходятся с Планом (П7): только в Акте — {sorted(a - p) or '—'}, "
                        f"только в Плане — {sorted(p - a) or '—'}")
        return _ok(f"исключения ({', '.join(sorted(a)) or 'нет'}) совпадают с Планом (П7)")

    return ctx.each(F.akt_fields, check)


def _akt_compliance(ctx: Ctx) -> dict:
    def check(f):
        sig, minor, text = f["nc_significant"], f["nc_minor"], (f["compliance"] or "").lower()
        if sig is None or minor is None:
            return _manual("число несоответствий в п.24 Акта не распознано")
        if not text:
            return _manual("формулировка п.14.4 Акта не найдена")
        if sig >= 1:
            expected, ok = "СМК частично соответствует критериям аудита", "частично" in text
        elif minor >= 1:
            expected, ok = "СМК в целом соответствует критериям аудита", "в целом" in text
        else:
            expected, ok = "СМК соответствует критериям аудита", ("частично" not in text and "в целом" not in text
                                                                 and "не соответствует" not in text)
        counts = f"значительных {sig}, малозначительных {minor}"
        if ok:
            return _ok(f"п.14.4 «{f['compliance'][:80]}» соответствует п.24 ({counts})")
        return _nok(f"п.14.4 «{f['compliance'][:80]}», а по п.24 ({counts}) должно быть «{expected}»")

    return ctx.each(F.akt_fields, check)


def _akt_pkd_deadline(ctx: Ctx) -> dict:
    def check(f):
        if not f["date"]:
            return _manual("дата Акта (Акт2) не распознана")
        if not f["pkd_deadline"]:
            return _manual("срок представления плана КД в п.17 Акта не найден")
        expected = F.workday_deadline(f["date"], 5)
        if f["pkd_deadline"] != expected:
            return _nok(f"срок в п.17 — {F.fmt_date(f['pkd_deadline'])}, а 5 р.д. с даты Акта "
                        f"({F.fmt_date(f['date'])}, день подписания — первый) — до {F.fmt_date(expected)}")
        return _ok(f"срок в п.17 — {F.fmt_date(expected)}: 5 рабочих дней с даты Акта "
                   f"(день подписания — первый)")

    return ctx.each(F.akt_fields, check)


def _akt_p24_consistency(ctx: Ctx) -> dict:
    # Формулировка подстроки шаблона «в соответствии с п.24» неоднозначна (вопрос методологу,
    # docs/ЧЕК-ЛИСТЫ.md §5) — даём оператору данные для ручной сверки.
    def check(f):
        return _manual(
            f"сверьте п.17 с результатами п.24: значительных {f['nc_significant']}, "
            f"малозначительных {f['nc_minor']}; п.14.4 — «{(f['compliance'] or '—')[:80]}»; "
            f"обязательство разработать план КД в п.17 — {'есть' if f['obligation_pkd'] else 'нет'}"
        )

    return ctx.each(F.akt_fields, check)


def _akt_kd_deadline(ctx: Ctx) -> dict:
    sert_name, sert = ctx.source("sertifikat", F.sert_fields)

    def check(f):
        main = f["types"] - {F.EXTENSION}
        if not f["date"]:
            return _manual("дата Акта (Акт2) не распознана")
        if not f["kd_deadline"]:
            return _manual("срок представления отчета в п.18 Акта (Акт6) не найден")
        if main & F.INSPECTION_TYPES:
            days, kind = 20, "ИК"
        elif main & {"первый этап", "второй этап", "ресертификация"}:
            days, kind = 40, "ресертификация" if "ресертификация" in main else "сертификация"
        else:
            return _manual(f"вид аудита в п.1 Акта не определён ({_types_label(f['types'])})")
        expected = F.workday_deadline(f["date"], days)
        if f["kd_deadline"] != expected:
            return _nok(f"срок в п.18 — {F.fmt_date(f['kd_deadline'])}, а для «{kind}» {days} р.д. с даты "
                        f"Акта ({F.fmt_date(f['date'])}, день подписания — первый) — до {F.fmt_date(expected)}")
        reason = f"срок в п.18 — {F.fmt_date(expected)}: {days} р.д. с даты Акта ({kind})"
        if kind != "ресертификация":
            return _ok(reason)
        if not sert_name or not sert["valid_until"]:
            return _manual(reason + "; срок действия сертификата (СЕРТ2) не найден — условие "
                                    "«≤ СЕРТ2 − 10 р.д.» проверьте вручную")
        limit = F.add_workdays(sert["valid_until"], -10)
        if f["kd_deadline"] > limit:
            return _nok(reason + f", но позже СЕРТ2 − 10 р.д. ({F.fmt_date(limit)}; сертификат до "
                                 f"{F.fmt_date(sert['valid_until'])})")
        return _ok(reason + f" и не позднее СЕРТ2 − 10 р.д. ({F.fmt_date(limit)})")

    return ctx.each(F.akt_fields, check)


def _akt_23_2(ctx: Ctx) -> dict:
    def check(f):
        if f["nc_23_2"] is None:
            return _manual("число несоответствий в п.23.2 Акта не найдено")
        if f["nc_significant"] is None or f["nc_minor"] is None:
            return _manual("число несоответствий в п.24 Акта не распознано")
        total = f["nc_significant"] + f["nc_minor"]
        if f["nc_23_2"] != total:
            return _nok(f"в п.23.2 — {f['nc_23_2']}, а в п.24: {f['nc_significant']} + {f['nc_minor']} = {total}")
        return _ok(f"в п.23.2 — {total} = п.24 ({f['nc_significant']} значит. + {f['nc_minor']} малознач.)")

    return ctx.each(F.akt_fields, check)


def _akt_distribution(ctx: Ctx) -> dict:
    return ctx.each(F.akt_fields, lambda f: _org(
        f["distribution_org"], f["org"], "наименование в Акте (А20)") if f["distribution_org"]
        else _manual("организация в списке рассылки (п.29) не найдена"))


def _akt_signatures(ctx: Ctx) -> dict:
    def check(f):
        members = {p for ps in f["team"].values() for p in ps}
        if not members:
            return _manual("состав ЭГ в Акте (А21) не распознан")
        if not f["signers"]:
            return _manual("блок подписей членов ЭГ не распознан")
        missing = sorted(members - set(f["signers"]))
        if missing:
            return _nok("в блоке подписей нет: " + ", ".join(map(F.person_label, missing)))
        return _ok("в блоке подписей все члены ЭГ: " + ", ".join(map(F.person_label, sorted(members)))
                   + " (собственноручные подписи на скане проверяются вручную)")

    return ctx.each(F.akt_fields, check)


def _akt_date(ctx: Ctx) -> dict:
    def check(f):
        if not f["date"] or not f["period"]:
            return _manual(("дата Акта (Акт2)" if not f["date"] else "сроки аудита (А8)") + " не распознаны")
        if f["date"] != f["period"][1]:
            return _nok(f"дата Акта {F.fmt_date(f['date'])}, а последний день аудита — {F.fmt_date(f['period'][1])}")
        return _ok(f"дата Акта {F.fmt_date(f['date'])} — последний день аудита")

    return ctx.each(F.akt_fields, check)


# ---------- Заключение об устранении несоответствий ----------

def _zakl_org(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return ctx.each(F.zakl_fields, lambda f: _org(f["org"], akt["org"], "Акт (Акт1)"))


def _zakl_persons_check(f) -> dict:
    if not f["approver"] or not f["leader"]:
        return _manual(("утверждающий" if not f["approver"] else "руководитель ЭГ") + " не распознан")
    if f["approver"] == f["leader"]:
        return _nok(f"Заключение утвердил руководитель ЭГ ({F.person_label(f['leader'])}) — "
                    f"утверждать должно другое лицо")
    return _ok(f"утвердил {F.person_label(f['approver'])}, руководитель ЭГ — {F.person_label(f['leader'])}")


def _zakl_persons(ctx: Ctx) -> dict:
    return ctx.each(F.zakl_fields, _zakl_persons_check)


def _zakl_date(ctx: Ctx) -> dict:
    kd_name, kd = ctx.source("otchet_kd", F.approved_doc_fields)
    if not kd_name:
        return _nok("Не загружен Отчет об устранении несоответствий (Отчет по КД) — сверять не с чем.")

    def check(f):
        if not f["approval_date"]:
            return _manual("дата утверждения Заключения не распознана")
        if f["sign_date"] and f["sign_date"] != f["approval_date"]:
            return _nok(f"дата утверждения {F.fmt_date(f['approval_date'])} ≠ дате подписания "
                        f"{F.fmt_date(f['sign_date'])}")
        if not kd["date"]:
            return _manual(f"дата утверждения Отчета по КД («{kd_name}») не распознана")
        limit = _plus_workdays(kd["date"], 5)
        if f["approval_date"] > limit:
            return _nok(f"утверждено {F.fmt_date(f['approval_date'])}, позже Отчета по КД "
                        f"({F.fmt_date(kd['date'])}) + 5 р.д. = {F.fmt_date(limit)}")
        return _ok(f"утверждено {F.fmt_date(f['approval_date'])} — не позднее Отчета по КД "
                   f"({F.fmt_date(kd['date'])}) + 5 р.д. ({F.fmt_date(limit)})")

    return ctx.each(F.zakl_fields, check)


def _zakl_types(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)

    def check(f):
        if not akt["types"]:
            return _manual("вид аудита в п.1 Акта (Акт7) не распознан")
        if not f["types"]:
            return _nok(f"в п.1 Заключения не отмечен вид аудита (в Акте: {_types_label(akt['types'])})")
        if f["types"] != akt["types"]:
            return _nok(f"в п.1 отмечено: {_types_label(f['types'])}; в Акте: {_types_label(akt['types'])}")
        return _ok(f"вид аудита совпадает с Актом: {_types_label(f['types'])}")

    return ctx.each(F.zakl_fields, check)


def _zakl_period(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)

    def check(f):
        if not f["period"] or not akt["period"]:
            return _manual("сроки проведения аудита в " + ("Заключении" if not f["period"] else "Акте (А8)")
                           + " не распознаны")
        if f["period"] != akt["period"]:
            return _nok(f"сроки {F.fmt_period(f['period'])} ≠ Акт: {F.fmt_period(akt['period'])}")
        return _ok(f"сроки {F.fmt_period(f['period'])} совпадают с Актом")

    return ctx.each(F.zakl_fields, check)


def _zakl_rows(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return ctx.each(F.zakl_fields, lambda f: _rows(f["rows"], akt["nc_total"]))


def _zakl_mark(key: str, label: str) -> Callable:
    def rule(ctx: Ctx) -> dict:
        def check(f):
            if f[key] is None:
                return _manual(f"отметка «{label}» в Заключении не найдена")
            return _ok(f"отмечено «{label}»") if f[key] else _nok(f"не отмечено «{label}»")
        return ctx.each(F.zakl_fields, check)
    return rule


def _zakl_decisions(ctx: Ctx) -> dict:
    """п.4 Заключения по подписям отметок, а не по кодам: коды в шаблоне (З10, З13, З14)
    и в образце (З13 — «Подтвердить», З12 — «Расширить») не совпадают."""
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)

    def check(f):
        types, note = f["types"], ""
        if not types:
            types, note = akt["types"], " (вид аудита взят из Акта: в п.1 Заключения не отмечен)"
        main = types - {F.EXTENSION}
        if not main:
            return _manual("вид аудита не определён ни в Заключении, ни в Акте")
        need = []
        if main & {"второй этап", "ресертификация"}:
            need.append(("issue", True, "«Выдать сертификат»"))
        if main & F.INSPECTION_TYPES:
            need.append(("confirm", True, "«Подтвердить действие сертификата»"))
        if F.EXTENSION in types or akt["extend"] or akt["narrow"]:
            # отметки «Расширить / Сузить область» должны совпадать с рекомендациями Акта (А40 / А41)
            need.append(("extend", bool(akt["extend"]), "«Расширить область сертификации»"))
            need.append(("narrow", bool(akt["narrow"]), "«Сузить область сертификации»"))
        problems = []
        for key, expected, label in need:
            actual = bool(f[key])
            if actual != expected:
                problems.append(f"{label} {'не отмечено' if expected else 'отмечено, но в Акте нет'}")
        if problems:
            return _nok("; ".join(problems) + note)
        shown = ", ".join(label for key, expected, label in need if expected)
        return _ok(f"отмечено: {shown}" + note)

    return ctx.each(F.zakl_fields, check)


def _zakl_distribution(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    ref = akt["distribution_org"] or akt["org"]
    return ctx.each(F.zakl_fields, lambda f: _org(
        f["distribution_org"], ref, "список рассылки Акта (Акт9)") if f["distribution_org"]
        else _manual("организация в списке рассылки Заключения не найдена"))


def _zakl_leader(ctx: Ctx) -> dict:
    akt_file, akt = ctx.akt
    if not akt_file:
        return _nok(NO_AKT)
    return ctx.each(F.zakl_fields, lambda f: _same_person(
        f["leader"], _akt_leader(akt), "Руководитель ЭГ", "Акт (Акт3)"))


# ---------- 1 этап: Отчет по результатам 1 этапа ----------

def _plus_workdays(d, n: int):
    """«Не позднее даты + n р.д.» вне сроков Акта (п.2 Отчета 1 этапа, п.3 Заключений): пока
    считается со следующего дня — распространять ли на них правило сроков Акта («день
    подписания — первый», F.workday_deadline), не решено (docs/ЧЕК-ЛИСТЫ.md §5)."""
    return F.add_workdays(d, n)


def _contract_refs(ctx: Ctx):
    """Реквизиты Д1: договор и все загруженные доп. соглашения. (номера, даты, файлы)."""
    nums, dates, files = set(), set(), []
    for dtype in ("dogovor", "dop_soglashenie"):
        for name in doc_types.files_of_type(ctx.doc_type_map, dtype):
            n, d = F.contract_fields(ctx.all_texts[name])["refs"]
            nums |= n
            dates |= d
            files.append(name)
    return nums, dates, files


def _o1_org(ctx: Ctx) -> dict:
    _, contract, err = _need(ctx, ("dogovor", "dop_soglashenie"), F.contract_fields, "Договор")
    if err:
        return err
    return ctx.each(F.otchet1_fields, lambda f: _org(f["org"], contract["org"], "Договор (Д2)", require_full=True))


def _o1_approval(ctx: Ctx) -> dict:
    def check(f):
        if not f["approval_date"]:
            return _manual("дата утверждения в шапке (1О7) не распознана")
        if not f["period"]:
            return _manual("сроки проведения аудита (1О8) не распознаны")
        limit = _plus_workdays(f["period"][1], 5)
        if f["approval_date"] > limit:
            return _nok(f"утверждён {F.fmt_date(f['approval_date'])}, позже срока аудита "
                        f"({F.fmt_period(f['period'])}) + 5 р.д. = {F.fmt_date(limit)}")
        return _ok(f"утверждён {F.fmt_date(f['approval_date'])} — не позднее срока аудита "
                   f"({F.fmt_period(f['period'])}) + 5 р.д. ({F.fmt_date(limit)})")

    return ctx.each(F.otchet1_fields, check)


def _o1_area(ctx: Ctx) -> dict:
    _, razb, err = _need(ctx, "razbivka", F.razbivka_fields, "файл «Разбивка области по кодам ОКВЭД»")
    if err:
        return err

    def check(f):
        words, ref = F.word_set(f["area"]), F.word_set(razb["area"])
        if not words:
            return _manual("область применения (1О5) в Отчете не распознана")
        if not ref:
            return _manual("заявленная область (ОБ1) в Разбивке не распознана")
        jac = F.jaccard(words, ref)
        reason = f"формулировка области: Разбивка (ОБ1) — совпадение слов {int(jac * 100)}%"
        return _ok(reason) if jac >= 0.7 else _nok(reason) if jac < 0.5 else _manual(reason + " — сверьте формулировку")

    return ctx.each(F.otchet1_fields, check)


def _o1_basis(ctx: Ctx) -> dict:
    nums, dates, files = _contract_refs(ctx)
    if not files:
        return _nok("Не загружен Договор — сверять не с чем.")

    def check(f):
        b_nums, b_dates = f["basis"]
        if not b_nums and not b_dates:
            return _manual("реквизиты основания (1О6) в Отчете не распознаны")
        if not nums and not dates:
            return _manual("номер и дата в Договоре / доп. соглашении не распознаны")
        missing = sorted(nums - b_nums) + sorted(F.fmt_date(d) for d in dates - b_dates)
        if missing:
            return _nok("в Основании нет реквизитов Договора / доп. соглашений: " + ", ".join(missing))
        extra = sorted(b_nums - nums) + sorted(F.fmt_date(d) for d in b_dates - dates)
        if extra:
            return _manual("в Основании есть реквизиты, которых нет в загруженных документах: "
                           + ", ".join(extra) + " — не загружено доп. соглашение?")
        return _ok(f"номера ({', '.join(sorted(nums))}) и даты Основания совпадают с Договором"
                   + (" и доп. соглашениями" if len(files) > 1 else ""))

    return ctx.each(F.otchet1_fields, check)


def _o1_team(ctx: Ctx) -> dict:
    _, prikaz, err = _need(ctx, "prikaz", F.prikaz_fields, "Приказ о назначении ЭГ")
    if err:
        return err

    def check(f):
        members = {p for ps in f["team"].values() for p in ps}
        if not members:
            return _manual("состав ЭГ (1О9) в Отчете не распознан")
        if not prikaz["team"]:
            return _manual("состав ЭГ в Приказе (ПР2) не распознан")
        ref = set(prikaz["team"])
        if members != ref:
            return _nok(f"состав ЭГ расходится с Приказом (ПР2): только в Отчете — "
                        f"{', '.join(map(F.person_label, sorted(members - ref))) or '—'}; только в Приказе — "
                        f"{', '.join(map(F.person_label, sorted(ref - members))) or '—'}")
        return _ok("состав ЭГ совпадает с Приказом (ПР2): " + ", ".join(map(F.person_label, sorted(members))))

    return ctx.each(F.otchet1_fields, check)


def _o1_nc_count(ctx: Ctx) -> dict:
    def check(f):
        if f["nc_count"] is None:
            return _manual("количество несоответствий (1О2) в Отчете не найдено")
        if f["nc_rows"] is None:
            return _manual("таблица несоответствий п.10.1.8 (1О10) не распознана")
        if f["nc_count"] != f["nc_rows"]:
            return _nok(f"указано несоответствий: {f['nc_count']}, а строк в таблице п.10.1.8: {f['nc_rows']}")
        return _ok(f"указано несоответствий: {f['nc_count']} — столько же строк в таблице п.10.1.8")

    return ctx.each(F.otchet1_fields, check)


def _o1_conformity(ctx: Ctx) -> dict:
    def check(f):
        if f["nc_count"] is None:
            return _manual("количество несоответствий (1О2) в Отчете не найдено")
        text = (f["conformity"] or "").lower()
        if not text:
            return _manual("формулировка п.10.1.4 не найдена")
        if f["nc_count"] == 0:
            expected = "Документация СМК соответствует процессам и деятельности организации"
            ok = "соответствует" in text and "в целом" not in text and "не соответствует" not in text
        else:
            expected = "Документация СМК в целом соответствует процессам и деятельности организации"
            ok = "в целом соответствует" in text
        quote = f["conformity"].strip("«»\" ")[:90]
        if ok:
            return _ok(f"п.10.1.4 «{quote}» соответствует числу несоответствий ({f['nc_count']})")
        return _nok(f"п.10.1.4 «{quote}», а при {f['nc_count']} несоответствиях должно быть «{expected}»")

    return ctx.each(F.otchet1_fields, check)


# ---------- 1 этап: Заключение 1 этапа ----------

def _need_otchet1(ctx: Ctx):
    return _need(ctx, "otchet1", F.otchet1_fields, "Отчет по результатам 1 этапа")


def _z1_org(ctx: Ctx) -> dict:
    _, o1, err = _need_otchet1(ctx)
    if err:
        return err
    return ctx.each(F.zakl1_fields, lambda f: _org(f["org"], o1["org"], "Отчет 1 этапа (1О1)"))


def _z1_date(ctx: Ctx) -> dict:
    _, ustr, err = _need(ctx, "otchet_ustr1", F.approved_doc_fields, "Отчет об устранении несоответствий 1 этапа")
    if err:
        return err

    def check(f):
        signed = f["sign_date"] or f["approval_date"]
        if not signed:
            return _manual("дата подписания Заключения (1З4) не распознана")
        if f["sign_date"] and f["approval_date"] and f["sign_date"] != f["approval_date"]:
            return _nok(f"дата подписания {F.fmt_date(f['sign_date'])} ≠ дате утверждения "
                        f"{F.fmt_date(f['approval_date'])}")
        if not ustr["date"]:
            return _manual("дата утверждения Отчета об устранении (1Отчет1) не распознана")
        limit = _plus_workdays(ustr["date"], 5)
        if signed > limit:
            return _nok(f"подписано {F.fmt_date(signed)}, позже Отчета об устранении "
                        f"({F.fmt_date(ustr['date'])}) + 5 р.д. = {F.fmt_date(limit)}")
        return _ok(f"подписано {F.fmt_date(signed)} — не позднее Отчета об устранении "
                   f"({F.fmt_date(ustr['date'])}) + 5 р.д. ({F.fmt_date(limit)})")

    return ctx.each(F.zakl1_fields, check)


def _z1_distribution(ctx: Ctx) -> dict:
    return ctx.each(F.zakl1_fields, lambda f: _org(
        f["distribution_org"], f["org"], "наименование в Заключении (1Закл1)") if f["distribution_org"]
        else _manual("организация в списке рассылки (1Закл5) не найдена"))


def _z1_rows(ctx: Ctx) -> dict:
    _, o1, err = _need_otchet1(ctx)
    if err:
        return err

    def check(f):
        if f["rows"] is None:
            return _manual("таблица несоответствий в Заключении не распознана")
        if o1["nc_count"] is None:
            return _manual("количество несоответствий в Отчете 1 этапа (1О2) не найдено")
        if f["rows"] != o1["nc_count"]:
            return _nok(f"строк в таблице {f['rows']}, а в Отчете 1 этапа несоответствий: {o1['nc_count']}")
        return _ok(f"строк в таблице {f['rows']} — столько же несоответствий в Отчете 1 этапа (1О2)")

    return ctx.each(F.zakl1_fields, check)


def _z1_mark(key: str, label: str) -> Callable:
    def rule(ctx: Ctx) -> dict:
        def check(f):
            if f[key] is None:
                return _manual(f"отметка «{label}» в Заключении не найдена")
            return _ok(f"отмечено «{label}»") if f[key] else _nok(f"не отмечено «{label}»")
        return ctx.each(F.zakl1_fields, check)
    return rule


# ---------- Привязка пунктов шаблонов к правилам ----------
# (подстроки названия пункта, правило). Название пункта нормализуется: нижний регистр,
# «ё» → «е», без пунктуации, пробелы схлопнуты. Первое совпадение выигрывает.

RULES = {
    "pkd": [
        (("название организации",), lambda c: _approved_org(c, require_full=False)),
        (("дата утверждения",), _pkd_date),
        (("количество строк",), _approved_rows),
    ],
    "otchet_kd": [
        (("название организации",), lambda c: _approved_org(c, require_full=True)),
        (("дата утверждения",), _kd_date),
        (("количество строк",), _approved_rows),
    ],
    "pn": [
        (("название организации",), _pn_org),
        (("фио руководителя",), _pn_leader),
        (("пункты нормативного",), _pn_sto),
        (("количество протоколов",), _pn_count),
        (("версия формы",), _pn_version),
        (("дата",), _pn_date),
    ],
    "loo_org": [
        (("название организации",), _loo_org),
        (("дата проверки",), _loo_period),
        (("оцениваемый",), _loo_person),
        (("фио",), _loo_signer),
        (("отметки в таблице",), _loo_marks),
        (("дата",), _loo_date),
    ],
    "loo_exp": [
        (("название организации",), _loo_org),
        (("дата проверки",), _loo_period),
        (("отметка в окошке",), _loo_standard),
        (("оцениваемый",), _loo_person),
        (("отметки в таблице",), _loo_marks),
        (("дата",), _loo_date),
    ],
    "akt": [
        (("наименование проверяемой организации",), _akt_org),
        (("вид аудита",), _akt_type),
        (("область применения",), _akt_area),
        (("номер и дата",), _akt_basis),
        (("адреса",), _akt_sites),
        (("коды",), _akt_okved),
        (("даты проведения",), _akt_period),
        (("уполномоченный представитель",), _akt_representative),
        (("состав эг",), _akt_team),
        (("исключения",), _akt_exclusions),
        (("14 4",), _akt_compliance),
        (("f:п 24",), _akt_p24_consistency),
        (("п 18",), _akt_kd_deadline),
        (("п 17",), _akt_pkd_deadline),
        (("23 2",), _akt_23_2),
        (("список рассылки",), _akt_distribution),
        (("подписи",), _akt_signatures),
        (("дата",), _akt_date),
    ],
    "zakl": [
        (("название организации",), _zakl_org),
        (("фио",), _zakl_persons),
        (("дата утверждения",), _zakl_date),
        (("отметка вида аудита",), _zakl_types),
        (("сроки проведения",), _zakl_period),
        (("количество строк",), _zakl_rows),
        (("несоответствия устранены",), _zakl_mark("eliminated", "Несоответствия устранены")),
        (("корректирующие действия реализованы",),
         _zakl_mark("implemented", "Корректирующие действия реализованы")),
        (("отметок в п 4",), _zakl_decisions),
        (("список рассылки",), _zakl_distribution),
        (("руководитель экспертной группы",), _zakl_leader),
    ],
    "otchet1": [
        (("наименование проверяемой организации",), _o1_org),
        (("дата утверждения",), _o1_approval),
        (("область применения",), _o1_area),
        (("основание",), _o1_basis),
        (("состав экспертной группы",), _o1_team),
        (("количество несоответствий",), _o1_nc_count),
        (("10 1 4",), _o1_conformity),
    ],
    "zakl1": [
        (("название организации",), _z1_org),
        (("фио",), lambda c: c.each(F.zakl1_fields, _zakl_persons_check)),
        (("дата подписания",), _z1_date),
        (("список рассылки",), _z1_distribution),
        (("количество строк",), _z1_rows),
        (("несоответствия устранены",), _z1_mark("eliminated", "Несоответствия устранены")),
        (("проведение второго этапа",),
         _z1_mark("second_stage", "Проведение второго этапа ресертификационного аудита возможно")),
    ],
}

SUPPORTED = set(RULES)


def _norm_area(s: str) -> str:
    s = re.sub(r"[^0-9a-zа-я ]", " ", (s or "").lower().replace("ё", "е"))
    return re.sub(r"\s+", " ", s).strip()


def rule_for(ctype: str, item: dict) -> Optional[Callable]:
    """Правило пункта: ключ — подстрока названия области, «f:…» — подстрока формулы."""
    area = _norm_area(item.get("area"))
    formula = _norm_area(item.get("comments"))
    for keys, rule in RULES.get(ctype, []):
        if any(k[2:] in formula if k.startswith("f:") else k in area for k in keys):
            return rule
    return None


def run_checklist(ctype: str, items: list, checked: dict, all_texts: dict,
                  doc_type_map: dict) -> dict:
    """Проверка чек-листа ctype. checked — {файл: текст} проверяемых документов.
    Результат в формате пайплайна: {"header": {}, "checklist": [{ok, nok, reason}]}."""
    results = []
    ctx = Ctx(checked, all_texts, doc_type_map)  # один на чек-лист: поля Акта разбираются один раз
    for item in items:
        rule = rule_for(ctype, item)
        ctx.formula = item.get("comments") or ""
        if rule is None:
            v = _manual("правило для пункта не задано")
        else:
            try:
                v = rule(ctx)
            except Exception as e:  # пункт не должен ронять весь чек-лист
                print(f"[formula {ctype}] ошибка в пункте «{item.get('area')}»: {e}")
                v = _manual(f"ошибка проверки: {str(e)[:120]}")
        reason = v["reason"]
        if v["status"] == "manual":
            reason = "ПРОВЕРИТЬ ВРУЧНУЮ: " + reason
        results.append({
            "ok": v["status"] == "ok",
            "nok": v["status"] == "nok",
            "reason": reason[:1500],
            "ii_data_found": "",
        })
        print(f"[formula {ctype}] {item.get('area', '')[:40]!r} → {v['status']}")
    return {"header": {}, "checklist": results}
