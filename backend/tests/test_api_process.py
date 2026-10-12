import json
import shutil

import pytest
from docx import Document
from fastapi.testclient import TestClient

import main

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


def _docx(path, lines):
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)


@pytest.fixture
def client(tmp_path, monkeypatch, checklist_path):
    uploads, outputs = tmp_path / "uploads", tmp_path / "outputs"
    sdir = uploads / SESSION
    sdir.mkdir(parents=True)
    outputs.mkdir()
    monkeypatch.setattr(main, "UPLOAD_DIR", uploads)
    monkeypatch.setattr(main, "OUTPUT_DIR", outputs)
    monkeypatch.setattr(main, "gigachat_preflight", lambda *a, **k: {"ok": True, "models": []})
    stub = json.dumps({"ok": False, "nok": True, "reason": "MODEL-STUB"}, ensure_ascii=False)
    monkeypatch.setattr(main, "_gigachat_call", lambda *a, **k: stub)

    for name in ("ИИ -План АУДИТА.docx", "Чек лист Акт.docx"):
        shutil.copy(checklist_path(name), sdir / name)
    _docx(sdir / "План аудита.docx", PLAN_LINES)
    _docx(sdir / "Приказ.docx", ["ПРИКАЗ № 1 13 июля 2026 г.", "О назначении экспертной группы",
                                 "Для проведения второго инспекционного контроля"])
    return TestClient(main.app)


def _process(client, template, template_2=None):
    data = {"api_key": "k", "template_file": template, "session_id": SESSION,
            "plan_doc_file": "План аудита.docx"}
    if template_2:
        data["template_file_2"] = template_2
    return client.post("/api/process", data=data)


def test_process_plan_2026(client):
    r = _process(client, "ИИ -План АУДИТА.docx")
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["status"] == "success"
    out = body["outputs"][0]
    assert out["type"] == "plan" and out["total_items"] == 12
    assert body["document_types"] == {"План аудита.docx": "plan", "Приказ.docx": "prikaz"}
    assert out["validation"]["checklist_total"] == 12


def test_unsupported_checklist_is_not_run_with_plan_rules(client):
    r = _process(client, "Чек лист Акт.docx")
    assert r.status_code == 400
    assert "«Акт»" in r.json()["detail"] and "этапе 2" in r.json()["detail"]


def test_unsupported_checklist_skipped_next_to_plan(client):
    r = _process(client, "Чек лист Акт.docx", "ИИ -План АУДИТА.docx")
    assert r.status_code == 200, r.text
    body = r.json()
    assert [o["type"] for o in body["outputs"]] == ["plan"]
    assert "Пропущены: «Акт»" in body["message"]
