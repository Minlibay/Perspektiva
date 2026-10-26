"""API пакета аудита (этап 4): загрузка по трём полям → разбор → фоновая проверка → статус."""
import io
import json
import time

import pytest
from docx import Document
from fastapi.testclient import TestClient

import main
import stage1_docs as S1
import stage2_docs as S2

SESSION = "ab" * 16

PLAN_LINES = [
    "ПЛАН аудита системы менеджмента качества",
    "УТВЕРЖДАЮ «17» июля 2026 г.",
    "1 Цель и область аудита",
    "5 Сроки проведения аудита 27-30.07.2026",
    "8 Исключения из требований стандартов п. 7.1.3.5 (абзац 1)",
    "9 Рабочий язык проведения аудита Русский",
    "1-й день. 27.07.2026",
    "(г. Пермь, ул. Ленина, д. 1)",
    "Предварительное совещание",
    "Заключительное совещание",
]
PRIKAZ_LINES = ["ПРИКАЗ № 1 13 июля 2026 г.", "О назначении экспертной группы",
                "Для проведения второго инспекционного контроля"]


def _docx_bytes(lines) -> bytes:
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


@pytest.fixture
def client(tmp_path, monkeypatch):
    uploads, outputs = tmp_path / "uploads", tmp_path / "outputs"
    (uploads / SESSION).mkdir(parents=True)
    outputs.mkdir()
    monkeypatch.setattr(main, "UPLOAD_DIR", uploads)
    monkeypatch.setattr(main, "OUTPUT_DIR", outputs)
    monkeypatch.setattr(main, "_JOBS", {})
    monkeypatch.setattr(main, "_ANALYSIS", {})
    monkeypatch.setattr(main, "_load_settings", lambda: {"api_key": "saved-key"})
    monkeypatch.setattr(main, "gigachat_preflight", lambda *a, **k: {"ok": True, "models": []})
    stub = json.dumps({"ok": False, "nok": True, "reason": "MODEL-STUB"}, ensure_ascii=False)
    monkeypatch.setattr(main, "_gigachat_call", lambda *a, **k: stub)
    return TestClient(main.app)


def upload(client, block: str, files: dict):
    """files — {имя: bytes}."""
    r = client.post("/api/upload", data={"session_id": SESSION, "block": block},
                    files=[("files", (name, data)) for name, data in files.items()])
    assert r.status_code == 200, r.text


def template(checklist_path, name) -> dict:
    return {name: checklist_path(name).read_bytes()}


def analyze(client, **body):
    r = client.post(f"/api/sessions/{SESSION}/analyze", json=body)
    assert r.status_code == 200, r.text
    return r.json()


def run_and_wait(client, checklists, **body):
    r = client.post(f"/api/sessions/{SESSION}/run", json={"checklists": checklists, **body})
    assert r.status_code == 200, r.text
    for _ in range(300):
        st = client.get(f"/api/sessions/{SESSION}/status").json()
        if not st["running"]:
            return st
        time.sleep(0.05)
    pytest.fail("проверка не завершилась")


def by_title(status) -> dict:
    return {o["title"]: o for o in status["outputs"]}


@pytest.fixture
def package(client, checklist_path):
    """ПКД и План в поле 1; План КД и План аудита в поле 2; Акт и Приказ в поле 3."""
    upload(client, "checklist", {**template(checklist_path, "Чек лист ПКД.docx"),
                                 **template(checklist_path, "ИИ -План АУДИТА.docx")})
    upload(client, "plan", {"План КД.docx": _docx_bytes(S1.PKD.splitlines()),
                            "План аудита.docx": _docx_bytes(PLAN_LINES)})
    upload(client, "sources", {"Акт.docx": _docx_bytes(S1.AKT.splitlines()),
                               "Приказ.docx": _docx_bytes(PRIKAZ_LINES)})
    return client


def test_analyze_three_fields(package):
    a = analyze(package)
    kinds = {f["name"]: (f["field"], f["kind"], f["doc_type"] or f["checklist_type"]) for f in a["files"]}
    assert kinds["Чек лист ПКД.docx"] == ("checklist", "checklist", "pkd")
    assert kinds["План КД.docx"] == ("plan", "document", "pkd")
    assert kinds["Акт.docx"] == ("sources", "document", "akt")
    cks = {c["title"]: c for c in a["checklists"]}
    assert cks["ПКД"]["selected"] == ["План КД.docx"] and not cks["ПКД"]["needs_gigachat"]
    assert cks["План АУДИТА"]["selected"] == ["План аудита.docx"] and cks["План АУДИТА"]["needs_gigachat"]
    assert all(c["ready"] for c in a["checklists"])


def test_run_formula_checklist_without_gigachat(package, monkeypatch):
    def no_gigachat(*a, **k):
        raise AssertionError("чек-лист «по формулам» не должен обращаться к GigaChat")

    monkeypatch.setattr(main, "gigachat_preflight", no_gigachat)
    monkeypatch.setattr(main, "_gigachat_call", no_gigachat)
    st = run_and_wait(package, ["Чек лист ПКД.docx"])
    assert st["stage"] == "done", st
    out = by_title(st)["ПКД"]
    assert (out["ok_count"], out["total_items"]) == (3, 3), out["checklist"]
    assert out["output_file"].endswith("Заполненный_ЧК_ПКД.docx")
    assert package.get(f"/api/download/{out['output_file']}").status_code == 200


def test_run_plan_and_pkd_together(package):
    st = run_and_wait(package, ["Чек лист ПКД.docx", "ИИ -План АУДИТА.docx"])
    outs = by_title(st)
    assert list(outs) == ["План АУДИТА", "ПКД"]           # План — первым
    assert outs["План АУДИТА"]["total_items"] == 12
    assert outs["План АУДИТА"]["validation"]["checklist_total"] == 12
    meta = json.loads((main.UPLOAD_DIR / SESSION / "_meta.json").read_text(encoding="utf-8"))
    assert [o["title"] for o in meta["outputs"]] == ["План АУДИТА", "ПКД"]
    assert "«Ромашка»" in meta["applicant"]                    # из проверяемого документа
    assert (meta["ok_count"], meta["total_items"]) == (outs["ПКД"]["ok_count"] + outs["План АУДИТА"]["ok_count"], 15)


def test_plan_needs_gigachat_key(package, monkeypatch):
    monkeypatch.setattr(main, "_load_settings", lambda: {})
    r = package.post(f"/api/sessions/{SESSION}/run", json={"checklists": ["ИИ -План АУДИТА.docx"]})
    assert r.status_code == 400 and "ключ GigaChat" in r.json()["detail"]
    # формульные чек-листы ключ не требуют
    assert package.post(f"/api/sessions/{SESSION}/run",
                        json={"checklists": ["Чек лист ПКД.docx"]}).status_code == 200


def test_gigachat_down_does_not_stop_other_checklists(package, monkeypatch):
    monkeypatch.setattr(main, "gigachat_preflight",
                        lambda *a, **k: {"ok": False, "stage": "oauth", "detail": "401"})
    outs = by_title(run_and_wait(package, ["ИИ -План АУДИТА.docx", "Чек лист ПКД.docx"]))
    assert "GigaChat недоступен" in outs["План АУДИТА"]["error"]
    assert outs["ПКД"]["ok_count"] == 3


def test_choose_between_two_akts(client, checklist_path):
    upload(client, "checklist", template(checklist_path, "Чек лист Акт.docx"))
    upload(client, "sources", {"Акт 2026.docx": _docx_bytes(S2.akt().splitlines()),
                               "Акт 2025.docx": _docx_bytes(S2.akt(okved="99.99").splitlines())})
    ck = analyze(client)["checklists"][0]
    assert ck["candidates"] == ["Акт 2025.docx", "Акт 2026.docx"]
    assert ck["selected"] == ["Акт 2025.docx"] and "несколько" in ck["warnings"][0]
    ck = analyze(client, choices={"Чек лист Акт.docx": ["Акт 2026.docx"]})["checklists"][0]
    assert ck["selected"] == ["Акт 2026.docx"]
    out = run_and_wait(client, ["Чек лист Акт.docx"],
                       choices={"Чек лист Акт.docx": ["Акт 2026.docx"]})["outputs"][0]
    assert out["checked_files"] == ["Акт 2026.docx"]
    assert "99.99" not in " ".join(it["reason"] for it in out["checklist"])


def test_overrides_and_misplaced_files(client, checklist_path):
    upload(client, "checklist", {**template(checklist_path, "Чек лист ПКД.docx"),
                                 "План аудита.docx": _docx_bytes(PLAN_LINES)})
    upload(client, "sources", {**template(checklist_path, "Чек лист ПН.docx"),
                               "скан.docx": _docx_bytes(["нечитаемый скан"]),
                               "Акт.docx": _docx_bytes(S1.AKT.splitlines())})
    a = analyze(client)
    warnings = " ".join(a["warnings"])
    assert "«План аудита.docx» в поле 1 не распознан как бланк" in warnings and "План аудита" in warnings
    assert "«Чек лист ПН.docx» (поле 3) — бланк чек-листа" in warnings
    assert "Тип документа «скан.docx» не определён" in warnings
    assert a["checklists"][0]["ready"] is False                # ПКД без проверяемого документа
    # пользователь указывает, что скан — это План КД
    a = analyze(client, overrides={"скан.docx": "pkd"})
    ck = a["checklists"][0]
    assert ck["ready"] and ck["selected"] == ["скан.docx"]
    scan = next(f for f in a["files"] if f["name"] == "скан.docx")
    assert scan["detected_by"] == "override"


def test_delete_file_and_reupload_to_other_field(package):
    r = package.delete(f"/api/files/План КД.docx?session_id={SESSION}")
    assert r.status_code == 200
    assert package.delete(f"/api/files/_meta.json?session_id={SESSION}").status_code == 404
    names = [f["name"] for f in analyze(package)["files"]]
    assert "План КД.docx" not in names
    # тот же файл, загруженный в другое поле, остаётся только в новом поле
    upload(package, "plan", {"Акт.docx": _docx_bytes(S1.AKT.splitlines())})
    akt = [f for f in analyze(package)["files"] if f["name"] == "Акт.docx"]
    assert [f["field"] for f in akt] == ["plan"]


def test_texts_are_cached(package, monkeypatch):
    analyze(package)
    calls = []
    monkeypatch.setattr(main, "_extract_text", lambda p: calls.append(p) or "")
    analyze(package)
    assert calls == []                                         # всё из кэша _texts.json


def test_status_restored_from_meta_after_restart(package, monkeypatch):
    run_and_wait(package, ["Чек лист ПКД.docx"])
    monkeypatch.setattr(main, "_JOBS", {})                      # «перезапуск сервера»
    st = package.get(f"/api/sessions/{SESSION}/status").json()
    assert st["stage"] == "done" and by_title(st)["ПКД"]["ok_count"] == 3


def test_bad_session(client):
    assert client.post("/api/sessions/zzz/analyze", json={}).status_code == 400
    assert client.get(f"/api/sessions/{'cd' * 16}/status").status_code == 404
