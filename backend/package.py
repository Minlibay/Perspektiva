"""
Разбор пакета аудита по трём полям загрузки (этап 4, docs/ЧЕК-ЛИСТЫ.md §7):
поле 1 — бланки чек-листов, поле 2 — проверяемые документы, поле 3 — источники.

Что определяется: тип каждого бланка, тип и организация каждого документа, какие
чек-листы можно запустить и с какими проверяемыми документами, и предупреждения
(нераспознанные файлы, чек-лист без документа, документы разных организаций…).
Модель не вызывается; тексты документов приходят уже извлечёнными.
"""

import hashlib
from typing import Optional

import checklists
import doc_types
import fields as F

FIELDS = ("checklist", "plan", "sources")
FIELD_TITLES = {"checklist": "поле 1", "plan": "поле 2", "sources": "поле 3"}
LLM_CHECKLISTS = {"plan", "svod"}           # проверяются через GigaChat
MULTI_DOC_CHECKLISTS = {"pn", "loo_org", "loo_exp"}   # проверяемых документов обычно несколько


def _detect_doc(name: str, text: Optional[str], override: Optional[str]) -> tuple:
    """(тип, как определён): override → по выбору пользователя, иначе по содержимому,
    иначе по имени файла."""
    if override:
        return override, "override"
    by_content = doc_types.classify_by_content(text or "")
    if by_content:
        return by_content, "content"
    by_name = doc_types.classify_by_name(name)
    return (by_name, "name") if by_name else ("unknown", None)


def _org_label(org: Optional[dict]) -> Optional[str]:
    return org["name"] if org else None


def analyze_package(files: list, overrides: Optional[dict] = None,
                    choices: Optional[dict] = None) -> dict:
    """files — [{"name", "field", "path", "text"}]; field ∈ FIELDS (файл без поля — источник).
    overrides — {файл: тип документа} от пользователя; choices — {бланк: [проверяемые файлы]}.

    Возвращает {"files", "checklists", "warnings", "doc_types"} — см. docs/ЧЕК-ЛИСТЫ.md §7.
    """
    overrides = overrides or {}
    choices = choices or {}
    out_files, warnings = [], []

    for f in files:
        field = f.get("field") if f.get("field") in FIELDS else "sources"
        path = f["path"]
        is_word = str(path).lower().endswith((".docx", ".docm"))
        ctype = checklists.detect_checklist_type(str(path)) if is_word else "unknown"
        entry = {"name": f["name"], "field": field, "kind": "document", "checklist_type": None,
                 "checklist_title": None, "items": None, "doc_type": None, "doc_title": None,
                 "detected_by": None, "org": None, "org_core": None, "org_opf": None, "used_by": []}

        if field == "checklist":
            if ctype == "unknown":
                entry["kind"] = "unknown"
                dtype, _ = _detect_doc(f["name"], f.get("text"), None)
                hint = (f" — похоже на документ «{doc_types.DOC_TYPE_BY_KEY[dtype]['title']}», "
                        f"перенесите его в поле 2 или 3" if dtype != "unknown" else "")
                warnings.append(f"«{f['name']}» в поле 1 не распознан как бланк чек-листа{hint}.")
            else:
                meta = checklists.CHECKLIST_BY_KEY[ctype]
                entry.update(kind="checklist", checklist_type=ctype, checklist_title=meta["title"],
                             items=len(checklists.parse_checklist(str(path))))
            out_files.append(entry)
            continue

        if ctype != "unknown":
            # бланк чек-листа не в своём поле — источником он быть не может
            entry.update(kind="misplaced_checklist", checklist_type=ctype,
                         checklist_title=checklists.CHECKLIST_BY_KEY[ctype]["title"])
            warnings.append(f"«{f['name']}» ({FIELD_TITLES[field]}) — бланк чек-листа "
                            f"«{entry['checklist_title']}». Перенесите его в поле 1; источником он не используется.")
            out_files.append(entry)
            continue

        dtype, how = _detect_doc(f["name"], f.get("text"), overrides.get(f["name"]))
        org = F.find_org(f.get("text") or "")
        entry.update(doc_type=dtype, detected_by=how,
                     doc_title=doc_types.DOC_TYPE_BY_KEY[dtype]["title"] if dtype in doc_types.DOC_TYPE_BY_KEY
                     else "тип не определён",
                     org=_org_label(org), org_core=org["key"][1] if org else None,
                     org_opf=org["key"][0] if org else None)
        if dtype == "unknown":
            warnings.append(f"Тип документа «{f['name']}» не определён — выберите его вручную, "
                            f"иначе документ не попадёт в проверку.")
        out_files.append(entry)

    docs = [e for e in out_files if e["kind"] == "document"]

    # --- чек-листы и их проверяемые документы ---
    out_checklists = []
    for e in out_files:
        if e["kind"] != "checklist":
            continue
        meta = checklists.CHECKLIST_BY_KEY[e["checklist_type"]]
        target = meta["checked_doc"]
        f2 = [d["name"] for d in docs if d["doc_type"] == target and d["field"] == "plan"]
        f3 = [d["name"] for d in docs if d["doc_type"] == target and d["field"] != "plan"]
        multi = e["checklist_type"] in MULTI_DOC_CHECKLISTS
        candidates = f2 + f3 if multi else (f2 or f3)
        chosen = [n for n in choices.get(e["name"], []) if n in candidates]
        selected = chosen or (candidates if multi else candidates[:1])
        ck = {"file": e["name"], "type": e["checklist_type"], "title": meta["title"], "items": e["items"],
              "needs_gigachat": e["checklist_type"] in LLM_CHECKLISTS,
              "checked_doc_type": target, "checked_doc_title": doc_types.DOC_TYPE_BY_KEY[target]["title"],
              "candidates": candidates, "selected": selected, "multi": multi,
              "ready": bool(selected), "problem": None, "warnings": []}
        if not selected:
            ck["problem"] = (f"нет проверяемого документа «{ck['checked_doc_title']}» — "
                             f"загрузите его в поле 2")
            warnings.append(f"Для чек-листа «{meta['title']}» нет проверяемого документа "
                            f"«{ck['checked_doc_title']}».")
        elif not multi and len(candidates) > 1:
            ck["warnings"].append(f"подходит несколько документов: {', '.join(candidates)} — "
                                  f"проверяется «{selected[0]}», можно выбрать другой")
        for name in selected:
            next(d for d in docs if d["name"] == name)["used_by"].append(e["name"])

        # другая организация у источников
        checked_orgs = {d["org_core"] for d in docs if d["name"] in selected and d["org_core"]}
        if checked_orgs:
            foreign = [d for d in docs if d["name"] not in selected and d["org_core"]
                       and d["org_core"] not in checked_orgs]
            if foreign:
                checked_label = next(d["org"] for d in docs if d["name"] in selected and d["org_core"])
                counts = {}   # (название, ОПФ) → [подпись, число файлов]: «ООО» и полная форма — одно
                for d in foreign:
                    counts.setdefault((d["org_core"], d["org_opf"]), [d["org"], 0])[1] += 1
                ck["warnings"].append(
                    f"проверяемый документ — {checked_label}; среди источников документы других "
                    f"организаций: " + "; ".join(f"{org} — {n} файл(ов)" for org, n in counts.values()))
        out_checklists.append(ck)

    # --- поле 2: документ, который не проверяется ни одним чек-листом ---
    for d in docs:
        if d["field"] == "plan" and not d["used_by"] and d["doc_type"] != "unknown":
            warnings.append(f"«{d['name']}» (поле 2, {d['doc_title']}) не проверяется ни одним "
                            f"загруженным чек-листом — нет подходящего бланка в поле 1.")

    # --- организации в пакете: группировка по названию; «ООО» и «Общество с ограниченной
    # ответственностью» — одна ОПФ, расхождение — только в нормализованной ОПФ (АО vs ООО)
    by_core = {}
    for d in docs:
        if d["org_core"]:
            by_core.setdefault(d["org_core"], []).append(d)

    def _group_label(group: list) -> str:
        by_opf = {}
        for d in group:
            by_opf.setdefault(d["org_opf"], d["org"])
        names = list(by_opf.values())
        return names[0] if len(names) == 1 else f"{' / '.join(names)} — ОПФ различается"

    if len(by_core) > 1:
        parts = [f"{_group_label(g)}: {len(g)} файл(ов)"
                 for g in sorted(by_core.values(), key=len, reverse=True)]
        warnings.append("В пакете документы нескольких организаций — " + "; ".join(parts) + ".")
    elif by_core:
        label = _group_label(next(iter(by_core.values())))
        if "ОПФ различается" in label:
            warnings.append(f"Организационно-правовая форма записана по-разному: {label.split(' — ')[0]}.")

    # --- дубликаты одного документа под разными именами ---
    by_hash = {}
    for f in files:
        text = (f.get("text") or "").strip()
        if len(text) > 200:
            by_hash.setdefault(hashlib.sha1(text.encode("utf-8")).hexdigest(), []).append(f["name"])
    for names in by_hash.values():
        if len(names) > 1:
            warnings.append(f"Один и тот же документ загружен несколько раз: {', '.join(names)}.")

    return {
        "files": out_files,
        "checklists": out_checklists,
        "warnings": warnings,
        "doc_types": [{"key": d["key"], "title": d["title"]} for d in doc_types.DOC_TYPES],
    }
