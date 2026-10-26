from docx import Document

import package
import stage1_docs as S1


def _f(tmp_path, name, field, text):
    path = tmp_path / name
    if name.endswith(".docx"):
        Document().save(path)       # пустой docx: не бланк чек-листа
    else:
        path.write_text(text, encoding="utf-8")
    return {"name": name, "field": field, "path": path, "text": text}


def test_organizations_full_and_short_names_are_one(tmp_path):
    files = [
        _f(tmp_path, "Акт.docx", "sources", S1.AKT),               # Общество с ограниченной … «Ромашка»
        _f(tmp_path, "Протоколы.docx", "sources", S1.pn()),       # ООО «Ромашка»
    ]
    assert package.analyze_package(files)["warnings"] == []


def test_several_organizations_and_opf_mismatch(tmp_path):
    files = [
        _f(tmp_path, "Акт.docx", "sources", S1.AKT),
        _f(tmp_path, "Сводный.docx", "plan", "СВОДНЫЙ АКТ\nАКЦИОНЕРНОЕ ОБЩЕСТВО «ПТС» (АО «ПТС»)"),
        _f(tmp_path, "Трудоемкость.docx", "sources", "РАСЧЕТ ТРУДОЕМКОСТИ\nНазвание Организации | ООО «ПТС»"),
    ]
    warnings = " ".join(package.analyze_package(files)["warnings"])
    assert "нескольких организаций" in warnings
    assert "АКЦИОНЕРНОЕ ОБЩЕСТВО «ПТС» / ООО «ПТС» — ОПФ различается" in warnings
    assert "Сводный.docx» (поле 2, Сводный акт) не проверяется ни одним" in warnings


def test_duplicates(tmp_path):
    files = [_f(tmp_path, "Акт.docx", "sources", S1.AKT), _f(tmp_path, "Акт (копия).docx", "sources", S1.AKT)]
    assert any("загружен несколько раз" in w for w in package.analyze_package(files)["warnings"])


def test_foreign_sources_flagged_per_checklist(tmp_path, checklist_path):
    files = [
        {"name": "Чек лист ПН.docx", "field": "checklist", "path": checklist_path("Чек лист ПН.docx"),
         "text": None},
        _f(tmp_path, "ПН.docx", "plan", S1.pn()),
        _f(tmp_path, "Акт другой.docx", "sources", S1.AKT.replace("«Ромашка»", "«Лютик»")),
    ]
    ck = package.analyze_package(files)["checklists"][0]
    assert ck["ready"] and ck["selected"] == ["ПН.docx"]
    assert "других организаций" in ck["warnings"][0] and "Лютик" in ck["warnings"][0]
