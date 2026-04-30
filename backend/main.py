"""
FastAPI backend для автоматического заполнения Плана АУДИТА
с использованием GigaChat API для извлечения данных из документов.
"""

import os
import sys
import shutil
import json
import uuid
import re

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')
from pathlib import Path
from typing import Optional
from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docx import Document
from docx.shared import Pt, Inches
import tempfile
from PyPDF2 import PdfReader
import openpyxl

app = FastAPI(title="Audit Plan Filler", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Папка для загруженных файлов
UPLOAD_DIR = Path(__file__).parent / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

_SESSION_ID_RE = re.compile(r"^[a-f0-9]{8,64}$")


def _resolve_session_dir(session_id: Optional[str]) -> tuple[str, Path]:
    """Возвращает (session_id, dir). Если id пустой — генерирует новый.
    Защита от path traversal: id должен быть hex.
    """
    if not session_id:
        session_id = uuid.uuid4().hex
    elif not _SESSION_ID_RE.match(session_id):
        raise HTTPException(status_code=400, detail="Некорректный session_id")
    sdir = UPLOAD_DIR / session_id
    sdir.mkdir(parents=True, exist_ok=True)
    return session_id, sdir

# Папка для результатов
OUTPUT_DIR = Path(__file__).parent / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

# Модель GigaChat по умолчанию. Из списка доступных на ключе:
#   "GigaChat" (Lite), "GigaChat-Pro", "GigaChat-Max",
#   "GigaChat-2", "GigaChat-2-Pro", "GigaChat-2-Max"  — новое поколение
# Реально используемая модель берётся из gigachat_settings.json (поле "model"),
# с фолбэком на это значение, если в настройках поле пустое.
DEFAULT_GIGACHAT_MODEL = "GigaChat-2-Pro"


def _load_settings() -> dict:
    """Прочитать gigachat_settings.json. Возвращает {} если нет или битый."""
    settings_file = Path(__file__).parent / "gigachat_settings.json"
    if not settings_file.exists():
        return {}
    try:
        with open(settings_file, encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def _get_active_model() -> str:
    """Активная модель: из настроек, иначе DEFAULT_GIGACHAT_MODEL."""
    return (_load_settings().get("model") or "").strip() or DEFAULT_GIGACHAT_MODEL

# Глобальный статус обработки — для опроса прогресса из UI
processing_status = {
    "stage": "idle",        # idle | header | verify | done | error
    "current": 0,           # текущий пункт чек-листа
    "total": 0,             # всего пунктов
    "message": "",          # текст для отображения
    "detail": "",           # подробность (файлы, этап)
}


class GigaChatSettings(BaseModel):
    """Настройки подключения к GigaChat"""
    api_key: str
    model: Optional[str] = None


class ProcessingResult(BaseModel):
    """Результат обработки"""
    status: str
    message: str
    extracted_data: Optional[dict] = None
    output_file: Optional[str] = None
    validation: Optional[dict] = None
    analyzed_files: Optional[list[str]] = None


_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"


def _extract_checkboxes_from_docx_xml(file_path: str) -> str:
    """Достаёт состояние галочек/крестиков из document.xml (SDT-checkbox и Wingdings-символы)."""
    import zipfile
    from xml.etree import ElementTree as ET

    findings = []
    try:
        with zipfile.ZipFile(file_path) as z:
            with z.open('word/document.xml') as f:
                tree = ET.parse(f)
    except Exception:
        return ""

    root = tree.getroot()

    # 1) SDT checkboxes (Word 2010+)
    for sdt in root.iter(f"{{{_W_NS}}}sdt"):
        cb = sdt.find(f".//{{{_W14_NS}}}checkbox")
        if cb is None:
            continue
        checked_el = cb.find(f"{{{_W14_NS}}}checked")
        is_checked = False
        if checked_el is not None:
            v = checked_el.get(f"{{{_W14_NS}}}val", "")
            is_checked = v in ("1", "true")
        ctx = ''.join(sdt.itertext()).strip()
        # пытаемся подцепить контекст параграфа
        parent_p = sdt
        while parent_p is not None and not parent_p.tag.endswith('}p'):
            parent_p = root.find(f".//*[.='{ctx}']/..") if False else None
            break
        findings.append((is_checked, ctx[:200]))

    # 2) Wingdings-символы галочек/крестиков в w:sym
    # Wingdings 2: F052 — ☒ (X в квадрате); F0A3 — V; F052 — крест.
    # Wingdings: F0FE — ☒, F0A8 — ☐.
    checked_codes = {"f0fe", "f052", "f0fc", "f0d8", "fe", "52", "fc"}
    unchecked_codes = {"f0a8", "a8", "f06f", "6f"}

    for para in root.iter(f"{{{_W_NS}}}p"):
        para_text = ''.join(para.itertext()).strip()
        for sym in para.iter(f"{{{_W_NS}}}sym"):
            char = sym.get(f"{{{_W_NS}}}char", "").lower()
            font = sym.get(f"{{{_W_NS}}}font", "").lower()
            if "wingdings" not in font and "symbol" not in font:
                continue
            if char in checked_codes:
                findings.append((True, para_text[:200]))
            elif char in unchecked_codes:
                findings.append((False, para_text[:200]))

    if not findings:
        return ""

    lines = ["=== ОТМЕТКИ (галочки/крестики) В ДОКУМЕНТЕ ==="]
    for checked, label in findings[:80]:
        mark = "☒ ОТМЕЧЕНО" if checked else "☐ НЕ ОТМЕЧЕНО"
        lines.append(f"{mark} | {label}")
    return "\n".join(lines)


_CHECKBOX_RE = re.compile(r"([☒☑✓✔✗✘×])\s*([^\n☒☑✓✔✗✘×]{1,150})")


def _scan_text_for_checkbox_symbols(text: str) -> str:
    """Сканирует текст на символы галочек/крестиков (для случаев когда они приходят как Unicode)."""
    if not text:
        return ""
    findings = []
    for m in _CHECKBOX_RE.finditer(text):
        sym = m.group(1)
        label = m.group(2).strip()
        if not label or len(label) < 2:
            continue
        is_checked = sym in "☒☑✓✔"
        mark = "ОТМЕЧЕНО" if is_checked else "НЕ ОТМЕЧЕНО"
        findings.append(f"{sym} {mark} | {label[:120]}")
    if not findings:
        return ""
    return "=== СИМВОЛЫ-ОТМЕТКИ В ТЕКСТЕ ===\n" + "\n".join(findings[:60])


def extract_text_from_docx(file_path: str) -> str:
    """Извлечение текста из .docx файла + распознанные галочки/крестики."""
    doc = Document(file_path)
    texts = []
    for para in doc.paragraphs:
        if para.text.strip():
            texts.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            row_texts = [cell.text.strip() for cell in row.cells]
            if any(row_texts):
                texts.append(" | ".join(row_texts))
    body = "\n".join(texts)

    extras = []
    cb_xml = _extract_checkboxes_from_docx_xml(file_path)
    if cb_xml:
        extras.append(cb_xml)
    cb_sym = _scan_text_for_checkbox_symbols(body)
    if cb_sym:
        extras.append(cb_sym)
    if extras:
        body = body + "\n\n" + "\n\n".join(extras)
    return body


_ocr_reader = None
_ocr_init_lock = __import__('threading').Lock()


def _get_ocr_reader():
    """Lazy-init EasyOCR (тяжёлая инициализация, делаем один раз). Потоко-безопасно."""
    global _ocr_reader
    if _ocr_reader is not None:
        return _ocr_reader
    with _ocr_init_lock:
        if _ocr_reader is not None:
            return _ocr_reader
        import easyocr
        prev_detail = processing_status.get("detail", "")
        processing_status["detail"] = "Инициализация OCR-движка (rus+eng)... первый запуск ~30-60 с"
        print("[ocr] Инициализация EasyOCR (rus+eng)... первый запуск качает модели ~200MB")
        _ocr_reader = easyocr.Reader(['ru', 'en'], gpu=False, verbose=False)
        print("[ocr] EasyOCR готов")
        processing_status["detail"] = prev_detail
    return _ocr_reader


def _ocr_pdf_pages(file_path: str, dpi: int = 200) -> str:
    """OCR всех страниц PDF через PyMuPDF (рендер) + EasyOCR (распознавание)."""
    import fitz
    reader = _get_ocr_reader()
    pages_text = []
    try:
        doc = fitz.open(file_path)
    except Exception as e:
        print(f"[ocr] Не удалось открыть {file_path}: {e}")
        return ""

    fname = Path(file_path).name
    page_count = doc.page_count
    try:
        for page_idx, page in enumerate(doc):
            try:
                processing_status["detail"] = f"OCR {fname}: стр {page_idx + 1}/{page_count}"
                # Рендерим страницу в PNG-байты
                pix = page.get_pixmap(dpi=dpi, alpha=False)
                img_bytes = pix.tobytes("png")
                # EasyOCR принимает байты
                results = reader.readtext(img_bytes, detail=0, paragraph=True)
                page_text = "\n".join(results) if results else ""
                if page_text.strip():
                    pages_text.append(f"--- Страница {page_idx + 1} (OCR) ---\n{page_text}")
                print(f"[ocr] {fname} стр.{page_idx+1}/{page_count}: {len(page_text)} симв.")
            except Exception as e:
                print(f"[ocr] Ошибка стр.{page_idx+1} {file_path}: {e}")
    finally:
        doc.close()

    return "\n\n".join(pages_text)


def _ocr_pdf_with_cache(file_path: str) -> str:
    """OCR с кэшированием в <file>.ocr.txt — повторно не запускаем."""
    cache = Path(file_path).with_suffix(Path(file_path).suffix + ".ocr.txt")
    if cache.exists():
        try:
            return cache.read_text(encoding="utf-8")
        except Exception:
            pass
    text = _ocr_pdf_pages(file_path)
    if text:
        try:
            cache.write_text(text, encoding="utf-8")
        except Exception as e:
            print(f"[ocr] Не удалось записать кэш {cache}: {e}")
    return text


def _extract_checkboxes_from_pdf_fields(file_path: str) -> str:
    """Достаёт значения чекбоксов из form fields PDF."""
    try:
        reader = PdfReader(file_path)
        fields = reader.get_fields() or {}
    except Exception:
        return ""
    lines = []
    for name, field in fields.items():
        try:
            ft = field.get('/FT')
            if ft != '/Btn':
                continue
            v = field.get('/V')
            v_str = str(v) if v is not None else ""
            is_checked = v_str not in ("", "/Off", "Off", "None")
            mark = "☒ ОТМЕЧЕНО" if is_checked else "☐ НЕ ОТМЕЧЕНО"
            lines.append(f"{mark} | поле '{name}'")
        except Exception:
            continue
    if not lines:
        return ""
    return "=== ЧЕКБОКСЫ PDF (form fields) ===\n" + "\n".join(lines[:80])


def extract_text_from_pdf(file_path: str) -> str:
    """Извлечение текста из .pdf файла + OCR (если текстовый слой пуст) + чекбоксы."""
    reader = PdfReader(file_path)
    page_count = len(reader.pages)
    texts = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            texts.append(text)
    body = "\n".join(texts)

    # Триггер OCR: если на страницу приходится меньше 100 символов — считаем что это скан
    avg_per_page = (len(body) / page_count) if page_count else 0
    if page_count > 0 and avg_per_page < 100:
        print(f"[ocr] {Path(file_path).name}: текстовый слой пуст ({avg_per_page:.0f} симв/стр), запускаю OCR...")
        try:
            ocr_text = _ocr_pdf_with_cache(file_path)
            if ocr_text:
                body = (body + "\n\n" if body else "") + "=== OCR-ТЕКСТ ===\n" + ocr_text
        except Exception as e:
            print(f"[ocr] Ошибка OCR {file_path}: {e}")

    extras = []
    cb_fields = _extract_checkboxes_from_pdf_fields(file_path)
    if cb_fields:
        extras.append(cb_fields)
    cb_sym = _scan_text_for_checkbox_symbols(body)
    if cb_sym:
        extras.append(cb_sym)
    if extras:
        body = body + "\n\n" + "\n\n".join(extras)
    return body


def extract_text_from_xlsx(file_path: str) -> str:
    """Извлечение текста из .xlsx файла"""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    texts = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        texts.append(f"=== Лист: {sheet_name} ===")
        for row in ws.iter_rows(values_only=True):
            row_vals = [str(cell) if cell is not None else "" for cell in row]
            if any(row_vals):
                texts.append(" | ".join(row_vals))
    return "\n".join(texts)


def fill_plan_template(template_path: str, source_texts: dict, 
                      extracted_data: dict, output_path: str) -> str:
    """Заполнение шаблона Плана АУДИТА извлечёнными данными"""
    shutil.copy2(template_path, output_path)
    
    doc = Document(output_path)
    
    # Проходим по всем таблицам документа
    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                for para in cell.paragraphs:
                    text = para.text.strip()
                    
                    # === ШАПКА (первая таблица, 4 строки) ===
                    if row_idx == 0 and col_idx == 0 and "Наименования Заявителя" in text:
                        # Правая ячейка — поле ввода
                        pass
                    elif "Наименования Заявителя" in text:
                        if "Наименование Заявителя" in extracted_data:
                            para.clear()
                            run = para.add_run(f"Наименования Заявителя: {extracted_data['Наименование Заявителя']}")
                            run.font.size = Pt(10)
                    elif "Вид аудита" in text and len(text) < 30:
                        if "Вид аудита" in extracted_data:
                            para.clear()
                            run = para.add_run(f"Вид аудита: {extracted_data['Вид аудита']}")
                            run.font.size = Pt(10)
                    elif "Даты проведения" in text:
                        if "Даты проведения" in extracted_data:
                            para.clear()
                            run = para.add_run(f"Даты проведения: {extracted_data['Даты проведения']}")
                            run.font.size = Pt(10)
                    elif text == "РЭГ" or (len(text) <= 5 and "РЭГ" in text):
                        if "РЭГ" in extracted_data:
                            para.clear()
                            run = para.add_run(f"РЭГ: {extracted_data['РЭГ']}")
                            run.font.size = Pt(10)
    
    doc.save(output_path)
    return output_path


def extract_checklist_from_template(template_path: str) -> list[dict]:
    """Извлечение структуры чек-листа из шаблона"""
    doc = Document(template_path)
    checklist = []
    
    for table in doc.tables:
        first_row = table.rows[0]
        cells_text = [cell.text.strip() for cell in first_row.cells]
        
        # Ищем таблицу чек-листа (5 колонок, заголовок "Область проверки")
        if len(first_row.cells) == 5 and "Область проверки" in cells_text:
            # Пропускаем заголовок, начинаем с row_idx=1
            for row_idx, row in enumerate(table.rows[1:], start=1):
                area = row.cells[0].text.strip()
                comments = row.cells[1].text.strip()
                problems_hint = row.cells[4].text.strip()
                
                # Извлекаем маркеры ИИ из комментариев
                import re
                ii_markers = re.findall(r'ИИ\d+', comments)
                
                if area:  # Пропускаем пустые строки
                    checklist.append({
                        "row_index": row_idx,
                        "area": area,
                        "comments": comments,
                        "problems_hint": problems_hint,
                        "ii_markers": ii_markers  # ['ИИ1', 'ИИ2'] и т.д.
                    })
            break  # Нашли таблицу чек-листа, больше не ищем
    
    return checklist


def extract_ii_references(template_path: str) -> dict:
    """
    Извлечение всех маркеров ИИ из шаблона и их контекста.
    Возвращает словарь: {"ИИ1": "текст рядом", "ИИ2": "текст рядом", ...}
    """
    import re
    doc = Document(template_path)
    ii_refs = {}
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                # Ищем все маркеры ИИ в ячейке
                markers = re.findall(r'ИИ\d+', text)
                for marker in markers:
                    # Запоминаем контекст — что написано рядом с маркером
                    # Берём предложение/фраза вокруг маркера
                    if marker not in ii_refs:
                        # Чистим текст от лишних пробелов и переносов
                        clean = re.sub(r'\s+', ' ', text)
                        ii_refs[marker] = clean
    
    return ii_refs


def fill_plan_with_checklist(template_path: str, extracted_data: dict, output_path: str, 
                              checklist_structure: list[dict] = None) -> str:
    """
    Заполнение шаблона Плана АУДИТА:
    - Шапка (Заявитель, Вид аудита, Даты, РЭГ)
    - Таблица чек-листа: ТОЛЬКО колонки ОК/NOK (текст НЕ трогаем)
    """
    shutil.copy2(template_path, output_path)
    
    doc = Document(output_path)
    
    checklist_data = extracted_data.get("checklist", [])
    header_data = extracted_data.get("header", {})
    
    # Счётчик заполненных строк чек-листа
    checklist_item_idx = 0
    
    for table in doc.tables:
        # Определяем тип таблицы
        first_row = table.rows[0]
        first_cell_text = first_row.cells[0].text.strip()
        cols_count = len(first_row.cells)
        
        # === ШАПКА (2 колонки) — заполняем ===
        if cols_count == 2:
            for row in table.rows:
                cells_text = [cell.text.strip() for cell in row.cells]
                full_row_text = " ".join(cells_text)
                
                if "Наименования Заявителя" in full_row_text:
                    if "Наименование Заявителя" in header_data:
                        for cell in row.cells:
                            if not cell.text.strip() or len(cell.text.strip()) < 5:
                                for para in cell.paragraphs:
                                    para.clear()
                                    run = para.add_run(header_data["Наименование Заявителя"])
                                    run.font.size = Pt(10)
                                break
                
                elif "Вид аудита" in full_row_text and len(full_row_text) < 50:
                    if "Вид аудита" in header_data:
                        for cell in row.cells:
                            if not cell.text.strip() or len(cell.text.strip()) < 5:
                                for para in cell.paragraphs:
                                    para.clear()
                                    run = para.add_run(header_data["Вид аудита"])
                                    run.font.size = Pt(10)
                                break
                
                elif "Даты проведения" in full_row_text:
                    if "Даты проведения" in header_data:
                        for cell in row.cells:
                            if not cell.text.strip() or len(cell.text.strip()) < 5:
                                for para in cell.paragraphs:
                                    para.clear()
                                    run = para.add_run(header_data["Даты проведения"])
                                    run.font.size = Pt(10)
                                break
                
                elif full_row_text.strip() == "РЭГ" or ("РЭГ" in full_row_text and len(full_row_text) < 20):
                    if "РЭГ" in header_data:
                        for cell in row.cells:
                            if not cell.text.strip() or len(cell.text.strip()) < 3:
                                for para in cell.paragraphs:
                                    para.clear()
                                    run = para.add_run(header_data["РЭГ"])
                                    run.font.size = Pt(10)
                                break
        
        # === ЧЕК-ЛИСТ (5 колонок, заголовок "Область проверки") ===
        elif cols_count == 5 and "Область проверки" in first_cell_text:
            # Пропускаем заголовок (row_idx=0), начинаем с row_idx=1
            for row in table.rows[1:]:
                if checklist_item_idx >= len(checklist_data):
                    break
                
                item = checklist_data[checklist_item_idx]
                
                # Колонки: 0=Область, 1=Комментарии, 2=ОК, 3=NOK, 4=Проблемные зоны
                # ВАЖНО: Трогаем ТОЛЬКО колонки 2 (ОК) и 3 (NOK)
                if len(row.cells) >= 5:
                    # === Колонка ОК — ставим ОДИН маркер ===
                    ok_cell = row.cells[2]
                    # Полная очистка ячейки — удаляем всё содержимое
                    for para in ok_cell.paragraphs:
                        for run in para.runs:
                            run.text = ""
                        para.text = ""
                    
                    # Если OK — записываем маркер
                    if item.get("ok"):
                        # Убедимся что есть хотя бы один параграф
                        if not ok_cell.paragraphs:
                            ok_cell.add_paragraph()
                        para = ok_cell.paragraphs[0]
                        run = para.add_run("☑")
                        run.font.size = Pt(11)
                    
                    # === Колонка NOK — ставим ОДИН маркер ===
                    nok_cell = row.cells[3]
                    # Полная очистка
                    for para in nok_cell.paragraphs:
                        for run in para.runs:
                            run.text = ""
                        para.text = ""
                    
                    # Если NOK — записываем маркер
                    if item.get("nok"):
                        if not nok_cell.paragraphs:
                            nok_cell.add_paragraph()
                        para = nok_cell.paragraphs[0]
                        run = para.add_run("☒")
                        run.font.size = Pt(11)
                    
                    # Колонки 0, 1, 4 НЕ ТРОГАЕМ — текст остаётся оригинальный из шаблона
                
                checklist_item_idx += 1
    
    doc.save(output_path)
    return output_path


def gigachat_preflight(api_key: str, timeout: float = 8.0, model: Optional[str] = None) -> dict:
    """
    Быстрая диагностика доступности GigaChat.
    Шаги:
      1. OAuth-запрос на порт 9443 (ngw.devices.sberbank.ru)
      2. Запрос списка моделей на порт 443 (gigachat.devices.sberbank.ru)
    Возвращает: {ok, stage, detail, models, current_model}.
    """
    import uuid
    active_model = (model or "").strip() or DEFAULT_GIGACHAT_MODEL
    try:
        import httpx
    except ImportError:
        return {"ok": False, "stage": "no_httpx", "detail": "httpx не установлен",
                "models": [], "current_model": active_model}

    result = {"ok": False, "stage": "", "detail": "", "models": [], "current_model": active_model}

    # Шаг 1: OAuth
    try:
        resp = httpx.post(
            "https://ngw.devices.sberbank.ru:9443/api/v2/oauth",
            headers={
                "Authorization": f"Basic {api_key}",
                "RqUID": str(uuid.uuid4()),
                "Content-Type": "application/x-www-form-urlencoded",
            },
            data={"scope": "GIGACHAT_API_PERS"},
            timeout=timeout,
            verify=False,
        )
    except httpx.ConnectTimeout:
        result["stage"] = "oauth_timeout"
        result["detail"] = ("Таймаут подключения к ngw.devices.sberbank.ru:9443. "
                            "Порт OAuth закрыт: firewall/антивирус/VPN/ISP. "
                            "Откройте исходящий TCP на 185.157.96.243:9443.")
        return result
    except httpx.ConnectError as e:
        result["stage"] = "oauth_connect"
        result["detail"] = f"Не удалось установить соединение с 9443: {str(e)[:150]}"
        return result
    except Exception as e:
        result["stage"] = "oauth_error"
        result["detail"] = f"{type(e).__name__}: {str(e)[:200]}"
        return result

    if resp.status_code == 401:
        result["stage"] = "oauth_unauthorized"
        result["detail"] = "Ключ отклонён (HTTP 401). Проверьте корректность авторизационных данных."
        return result
    if resp.status_code >= 400:
        result["stage"] = "oauth_http_error"
        result["detail"] = f"OAuth вернул HTTP {resp.status_code}: {resp.text[:200]}"
        return result

    try:
        access_token = resp.json().get("access_token")
    except Exception as e:
        result["stage"] = "oauth_parse"
        result["detail"] = f"Не удалось распарсить ответ OAuth: {str(e)[:120]}"
        return result

    if not access_token:
        result["stage"] = "oauth_no_token"
        result["detail"] = "OAuth ответил 200, но в ответе нет access_token"
        return result

    # Шаг 2: Список моделей
    try:
        models_resp = httpx.get(
            "https://gigachat.devices.sberbank.ru/api/v1/models",
            headers={"Authorization": f"Bearer {access_token}"},
            timeout=timeout,
            verify=False,
        )
    except Exception as e:
        result["stage"] = "models_request_error"
        result["detail"] = f"Ошибка запроса /models: {type(e).__name__}: {str(e)[:150]}"
        return result

    if models_resp.status_code != 200:
        result["stage"] = "models_http_error"
        result["detail"] = f"/models вернул HTTP {models_resp.status_code}: {models_resp.text[:200]}"
        return result

    try:
        result["models"] = [m.get("id") for m in models_resp.json().get("data", []) if m.get("id")]
    except Exception as e:
        result["stage"] = "models_parse"
        result["detail"] = f"Не удалось распарсить список моделей: {str(e)[:120]}"
        return result

    result["ok"] = True
    result["stage"] = "ready"
    if active_model in result["models"]:
        result["detail"] = f"OAuth и список моделей получены. Текущая модель '{active_model}' доступна."
    else:
        result["detail"] = (f"OAuth работает, но модель '{active_model}' отсутствует в списке доступных. "
                            f"Выберите одну из: {', '.join(result['models'])}.")
        result["ok"] = False
        result["stage"] = "model_not_available"
    return result


def _gigachat_call(api_key: str, system_prompt: str, user_prompt: str,
                   model: str = "GigaChat", temperature: float = 0.0,
                   max_tokens: int = 2000) -> str:
    """Один вызов GigaChat с retry до 3 попыток."""
    import time
    from gigachat import GigaChat

    last_err = None
    for attempt in range(3):
        try:
            gc = GigaChat(credentials=api_key, verify_ssl_certs=False)
            response = gc.chat({
                "model": model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": temperature,
                "max_tokens": max_tokens,
            })
            return response.choices[0].message.content
        except Exception as e:
            last_err = e
            print(f"[gigachat] Попытка {attempt+1} провалилась: {e}")
            time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"GigaChat недоступен после 3 попыток: {last_err}")


def _parse_json_response(text: str) -> dict:
    """Извлечь и распарсить JSON из ответа GigaChat (с поддержкой markdown-блоков)."""
    import re
    cleaned = text.strip()
    # Снимаем markdown-обёртку ```json ... ```
    cleaned = re.sub(r'^```(?:json)?\s*', '', cleaned)
    cleaned = re.sub(r'\s*```$', '', cleaned)

    start = cleaned.find('{')
    if start < 0:
        raise ValueError(f"JSON не найден: {cleaned[:200]}")

    depth = 0
    in_string = False
    escape = False
    for i in range(start, len(cleaned)):
        ch = cleaned[i]
        if escape:
            escape = False
            continue
        if ch == '\\':
            escape = True
            continue
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                return json.loads(cleaned[start:i+1])
    raise ValueError(f"Не удалось найти закрывающую скобку JSON: {cleaned[:300]}")


def extract_header_info(api_key: str, all_texts: dict, model: str = "GigaChat") -> dict:
    """Извлечь данные шапки (Заявитель, Вид аудита, Даты, РЭГ) из пакета документов."""
    # Строим компактный пакет: первые 3000 символов каждого файла
    summaries = []
    for fname, text in all_texts.items():
        summaries.append(f"=== {fname} ===\n{text[:3000]}")
    combined = "\n\n".join(summaries)[:80000]

    system_prompt = """Ты извлекаешь 4 поля для ШАПКИ Плана АУДИТА из документов.

Верни ТОЛЬКО JSON:
{
  "Наименование Заявителя": "полное название организации-заявителя на сертификацию СМК",
  "Вид аудита": "один из: 'Сертификационный аудит', 'Аудит', 'Расширение', 'Дополнительный аудит', 'Инспекционный контроль'",
  "Даты проведения": "даты в формате ДД.ММ.ГГГГ-ДД.ММ.ГГГГ или ДД.ММ.ГГГГ-ДД.ММ.ГГГГ",
  "РЭГ": "регистрационный номер дела (обычно начинается с номера типа 01-01-2025 или подобного)"
}

Если конкретное поле не найдено в документах — напиши "не найдено". НЕ выдумывай данные."""

    user_prompt = f"Документы:\n\n{combined}"
    response = _gigachat_call(api_key, system_prompt, user_prompt, model=model, max_tokens=800)
    return _parse_json_response(response)


def classify_files_for_item(api_key: str, item: dict, ii_references: dict,
                             file_names: list, model: str = "GigaChat") -> list:
    """
    Определить, в каких файлах искать данные для конкретного пункта чек-листа.
    Возвращает список имён файлов (подмножество file_names).
    """
    ii_markers = item.get("ii_markers", [])
    ii_context = "\n".join([f"- {m}: {ii_references.get(m, '')[:150]}" for m in ii_markers])

    system_prompt = """Ты определяешь, в каких файлах искать данные для одного пункта чек-листа аудита.

Имена файлов содержат описание содержимого (например "4.1. Приказ об ЭГ..."). Используй это.

Верни ТОЛЬКО JSON:
{
  "relevant_files": ["имя_файла_1", "имя_файла_2", ...]
}

ПРАВИЛА:
- Включи ВСЕ потенциально релевантные файлы (лучше лишний, чем пропущенный).
- Имена файлов должны быть точно такими же, как в списке ниже (копируй буква в букву).
- Максимум 20 файлов.
- Если не уверен какие — включи все файлы с совпадающим номером раздела (например, все файлы "4.x" для пункта раздела 4)."""

    files_list = "\n".join([f"- {f}" for f in file_names])

    user_prompt = f"""ПУНКТ ЧЕК-ЛИСТА:
ОБЛАСТЬ: {item['area']}
КОММЕНТАРИИ: {item['comments']}
ПОДСКАЗКА: {item['problems_hint']}
МАРКЕРЫ ИИ: {', '.join(ii_markers) if ii_markers else 'нет'}

СПРАВОЧНИК МАРКЕРОВ (что они означают):
{ii_context}

ДОСТУПНЫЕ ФАЙЛЫ:
{files_list}"""

    try:
        response = _gigachat_call(api_key, system_prompt, user_prompt, model=model, max_tokens=800)
        result = _parse_json_response(response)
        relevant = result.get("relevant_files", [])
        # Валидация имён (точное совпадение) + case-insensitive фолбэк
        valid = [f for f in relevant if f in file_names]
        if not valid:
            lower_map = {f.lower(): f for f in file_names}
            valid = [lower_map[r.lower()] for r in relevant if r.lower() in lower_map]
        return valid if valid else _fallback_files_by_section(item, file_names)
    except Exception as e:
        print(f"[classify] Ошибка ({item.get('area', '')[:30]}): {e}. Fallback по номеру раздела.")
        return _fallback_files_by_section(item, file_names)


def _section_prefix(s: str) -> str:
    """Достаёт ведущий номер раздела вида '4', '4.1', '4.1.2' из строки.
    Возвращает '' если не найден.
    """
    m = re.match(r'\s*(\d+(?:\.\d+)*)', s or '')
    return m.group(1) if m else ''


# Стоп-слова: служебные/общие, не несут различающего смысла для поиска по аудиту.
_RU_STOPWORDS = {
    "который", "которая", "которое", "которые", "если", "может", "должен", "должна",
    "должны", "должно", "также", "более", "менее", "очень", "когда", "пока",
    "соответствие", "соответствует", "соответствующий", "наличие", "иметь",
    "иметься", "проверка", "проверяется", "пункт", "пункта", "пункту", "область",
    "областью", "комментарий", "комментарии", "проблемы", "проблема", "зона",
    "зоны", "данные", "данных", "документ", "документа", "документы", "документов",
    "файл", "файла", "файлы", "файлов", "требования", "требование", "требуется",
    "только", "нужно", "необходимо", "следует", "часть", "части", "случае",
    "процесс", "процесса", "вопрос", "вопросы", "наличие", "необходимость",
    "осуществляется", "проведение", "проведения", "проводится", "выполнение",
    "выполнения", "выполняется", "включая", "относится", "связан", "связана",
    "перечень", "перечня", "форма", "формы", "образец", "образца",
}


def _extract_keywords(*texts: str) -> set:
    """Достаёт значимые слова (≥4 символов) из переданных строк, без стоп-слов."""
    out = set()
    for t in texts:
        for w in re.findall(r'[А-Яа-яЁёA-Za-z][А-Яа-яЁёA-Za-z\-]{3,}', t or ''):
            wl = w.lower()
            if wl in _RU_STOPWORDS:
                continue
            out.add(wl)
    return out


def find_relevant_files_for_item(item: dict, all_texts: dict) -> list:
    """
    Детерминированный подбор файлов под пункт чек-листа.
    Сигналы:
      A) Номер раздела пункта совпадает с префиксом имени файла,
         либо встречается отдельной "лексемой" в тексте файла.
      B) ≥2 ключевых слов из area/comments/problems_hint встречаются в тексте файла.
    Файл попадает в результат, если сработал хотя бы один сигнал.
    Если ничего не найдено — возвращает все файлы (лучше избыточно, чем пропустить).
    Возвращает список, отсортированный по убыванию релевантности.
    """
    item_section = _section_prefix(item.get('area', ''))
    item_top = item_section.split('.')[0] if item_section else ''
    section_re = re.compile(rf'(?<!\d){re.escape(item_section)}(?!\d)') if item_section else None

    keywords = _extract_keywords(
        item.get('area', ''), item.get('comments', ''), item.get('problems_hint', '')
    )

    scored = []
    for fname, text in all_texts.items():
        score = 0
        # Сигнал A: совпадение по номеру раздела
        fname_section = _section_prefix(fname)
        if item_section and fname_section:
            if fname_section == item_section:
                score += 3
            elif fname_section.startswith(item_section + '.') or item_section.startswith(fname_section + '.'):
                score += 2
            elif item_top and fname_section.split('.')[0] == item_top:
                score += 1
        # Номер раздела упомянут в самом тексте файла
        if section_re and text and section_re.search(text):
            score += 1

        # Сигнал B: совпадение ключевых слов в тексте
        if keywords and text:
            text_lower = text.lower()
            hits = sum(1 for kw in keywords if kw in text_lower)
            if hits >= 2:
                score += min(hits // 2, 3)  # ограничиваем вклад чтобы не перебивать раздел

        if score > 0:
            scored.append((fname, score))

    scored.sort(key=lambda x: -x[1])
    if scored:
        return [f for f, _ in scored]
    return list(all_texts.keys())


def _fallback_files_by_section(item: dict, file_names: list) -> list:
    """Фолбэк-выбор файлов: те, чей префикс номера совпадает с разделом пункта.
    Если совпадений нет — возвращает все файлы (лучше избыточно, чем пропустить).
    """
    item_section = _section_prefix(item.get('area', ''))
    if not item_section:
        return list(file_names)
    # Берём верхний уровень раздела ('4.1' -> '4', '4' -> '4')
    top = item_section.split('.')[0]
    matched = [f for f in file_names if _section_prefix(f).split('.')[0] == top]
    return matched if matched else list(file_names)


def build_evidence_pack(relevant_files: list, all_texts: dict, max_chars: int = 180000) -> str:
    """
    Собрать пак доказательств из релевантных файлов.
    Бюджет делится поровну между файлами, минимальный потолок снят —
    при большом числе файлов лучше короткий фрагмент из каждого, чем пропуск.
    """
    parts = []
    if not relevant_files:
        return ""
    per_file_budget = max(max_chars // len(relevant_files), 2000)
    for fname in relevant_files:
        if fname not in all_texts:
            continue
        text = all_texts[fname]
        chunk = text[:per_file_budget]
        parts.append(f"=== ФАЙЛ: {fname} ===\n{chunk}")
    return "\n\n".join(parts)


def verify_item_strict(api_key: str, item: dict, ii_references: dict,
                        evidence: str, model: str = "GigaChat",
                        extra_instructions: str = "") -> dict:
    """
    Строгая проверка пункта чек-листа с NOK-first логикой.
    Возвращает: {ok, nok, reason, ii_data_found, evidence_quote, source_file}.
    """
    ii_markers = item.get("ii_markers", [])
    ii_context = "\n".join([f"- {m}: {ii_references.get(m, '')[:150]}" for m in ii_markers])

    system_prompt = """Ты опытный аудитор СМК. Объективно проверяешь один пункт чек-листа аудита по документам-источникам.

ПРАВИЛА ОЦЕНКИ:
1. Оценивай ТОЛЬКО по ФАКТАМ из документов. Не делай выводов без прямых доказательств из текста.
2. OK — только при ЯВНОМ подтверждении: в файлах есть конкретный документ/факт, прямо соответствующий требованию пункта. Косвенные признаки и предположения — НЕ основание для OK.
3. NOK — когда:
   - требуемые документы/данные отсутствуют в загруженных файлах,
   - ИЛИ найдены существенные несоответствия (неправильная дата, неверный номер, устаревший документ, пропущенная подпись и т.п.),
   - ИЛИ есть противоречие между документом и требованием,
   - ИЛИ данных недостаточно для уверенного вывода о соответствии.
4. По умолчанию — NOK. OK только если есть чёткая цитата/факт, подтверждающий выполнение требования.
5. Если в пункте упомянуты маркеры ИИ (ИИ1, ИИ2...) — используй их как указатели, где искать. Но отсутствие самого слова "ИИ1" в файле не означает NOK: важны данные, а не маркер.
6. Не выдумывай факты. Если в предоставленных файлах данных нет — напиши это в reason и поставь NOK.

Верни ТОЛЬКО JSON (без markdown-обёртки, без комментариев):
{
  "ok": true/false,
  "nok": true/false,
  "evidence_quote": "цитата или пересказ из файла (до 350 символов), подтверждающая вердикт",
  "source_file": "имя файла где найдено доказательство, или 'не найдено'",
  "reason": "обоснование (1-3 предложения): что именно нашёл/не нашёл и почему такой вердикт",
  "ii_data_found": "список найденных маркеров через запятую (ИИ1, ИИ2...), или 'не найдено'"
}

ok и nok — взаимоисключающие: ровно один true, другой false."""

    user_prompt = f"""ПУНКТ ДЛЯ ПРОВЕРКИ:
ОБЛАСТЬ: {item['area']}
ТРЕБОВАНИЯ (КОММЕНТАРИИ В ШАБЛОНЕ): {item['comments']}
ПОДСКАЗКА (ПРОБЛЕМНЫЕ ЗОНЫ): {item['problems_hint']}
МАРКЕРЫ ИИ В ПУНКТЕ: {', '.join(ii_markers) if ii_markers else 'нет'}

СПРАВОЧНИК МАРКЕРОВ ИИ:
{ii_context if ii_context else '(нет маркеров)'}

ДОКУМЕНТЫ-ИСТОЧНИКИ ДЛЯ ПРОВЕРКИ:
{evidence if evidence else '(файлы не переданы)'}
{('\n\nДОПОЛНИТЕЛЬНЫЕ ПРАВИЛА ДЛЯ ЭТОГО ПУНКТА:\n' + extra_instructions) if extra_instructions else ''}

Проведи СТРОГУЮ проверку. Помни: по умолчанию NOK. OK только при явном доказательстве."""

    response = _gigachat_call(api_key, system_prompt, user_prompt, model=model, max_tokens=2500)
    result = _parse_json_response(response)

    # Нормализация: ровно один из ok/nok должен быть true.
    # Дефолт при неоднозначности — NOK (нет явного подтверждения = не соответствует).
    ok = bool(result.get("ok", False))
    nok = bool(result.get("nok", False))
    if ok == nok:
        reason_text = (result.get("reason", "") + " " + result.get("evidence_quote", "")).lower()
        ok_signals = ["подтвержд", "соответствует", "найден", "имеется", "присутств", "выполнен"]
        has_ok_signal = any(sig in reason_text for sig in ok_signals)
        # Перевешиваем в OK только если явный позитивный сигнал И нет негативных
        nok_signals = ["не найден", "отсутств", "не соответ", "противоречи", "не совпад", "нет данных", "не обнаруж", "не указан", "не подтвержд"]
        has_nok_signal = any(sig in reason_text for sig in nok_signals)
        if has_ok_signal and not has_nok_signal:
            ok, nok = True, False
        else:
            ok, nok = False, True
    result["ok"] = ok
    result["nok"] = nok
    result.setdefault("reason", result.get("evidence_quote", ""))
    result.setdefault("ii_data_found", ", ".join(ii_markers) if ii_markers else "нет")
    result.setdefault("evidence_quote", "")
    result.setdefault("source_file", "не найдено")
    return result


def adversarial_recheck(api_key: str, item: dict, ii_references: dict,
                         evidence: str, prior_verdict: dict,
                         model: str = "GigaChat") -> dict:
    """
    Второй проход для OK-вердиктов: попытаться найти причины перевести в NOK.
    Это основная защита от ложных OK.
    """
    ii_markers = item.get("ii_markers", [])

    system_prompt = """Ты — второй аудитор, который делает критический ревью вердикта OK от первого аудитора. Твоя задача — объективно подтвердить OK или перевернуть в NOK, если есть СУЩЕСТВЕННЫЕ основания.

СУЩЕСТВЕННЫЕ основания для NOK (только такие!):
1. Первый аудитор сослался на файл/цитату, которой на самом деле нет в документах-источниках ниже.
2. Явное противоречие между разными документами, которое первый пропустил.
3. Отсутствие критичного реквизита, без которого пункт не может считаться выполненным (например, нет подписи на приказе, нет даты на документе, не тот номер).
4. Данные относятся к ДРУГОМУ аудиту/заявителю/периоду.
5. Фальсификация цитаты первым аудитором.

Несущественные причины (НЕ переворачивать OK в NOK):
- "Мог бы быть более подробный" документ — не причина.
- "Первый аудитор сформулировал обоснование коротко" — не причина.
- "Не все маркеры ИИ подтверждены" — если найдены основные, этого достаточно.
- "Теоретически где-то мог бы быть нюанс" — спекуляция, не переворачивай.

ДЕФОЛТ: подтверждать OK. NOK — только при чётких существенных основаниях, которые ты можешь указать цитатой из файла.

Верни ТОЛЬКО JSON (та же схема):
{
  "ok": true/false,
  "nok": true/false,
  "evidence_quote": "...",
  "source_file": "...",
  "reason": "если OK — 'Подтверждено: ...'. Если NOK — укажи КОНКРЕТНОЕ существенное основание.",
  "ii_data_found": "..."
}"""

    user_prompt = f"""ПУНКТ:
ОБЛАСТЬ: {item['area']}
ТРЕБОВАНИЯ: {item['comments']}
ПОДСКАЗКА: {item['problems_hint']}
МАРКЕРЫ ИИ: {', '.join(ii_markers) if ii_markers else 'нет'}

ВЕРДИКТ ПЕРВОГО АУДИТОРА: OK
ЕГО ОБОСНОВАНИЕ: {prior_verdict.get('reason', '')}
ЦИТАТА, КОТОРУЮ ОН ПРИВЁЛ: {prior_verdict.get('evidence_quote', '')}
ФАЙЛ-ИСТОЧНИК: {prior_verdict.get('source_file', '')}

ДОКУМЕНТЫ-ИСТОЧНИКИ (проверь их ещё раз):
{evidence if evidence else '(нет)'}

Найди причины поменять на NOK. Если не нашёл — подтверди OK."""

    try:
        response = _gigachat_call(api_key, system_prompt, user_prompt, model=model, max_tokens=2500)
        result = _parse_json_response(response)

        ok = bool(result.get("ok", False))
        nok = bool(result.get("nok", False))
        if ok == nok:
            # Неоднозначный ответ ревью — оставляем исходный OK (не переворачиваем без явных оснований)
            ok, nok = True, False
        result["ok"] = ok
        result["nok"] = nok
        result.setdefault("reason", prior_verdict.get("reason", ""))
        result.setdefault("ii_data_found", prior_verdict.get("ii_data_found", ""))
        result.setdefault("evidence_quote", prior_verdict.get("evidence_quote", ""))
        result.setdefault("source_file", prior_verdict.get("source_file", "не найдено"))
        return result
    except Exception as e:
        print(f"[adversarial] Ошибка: {e}. Оставляю исходный вердикт OK.")
        # Если ревью не прошло (сетевая ошибка, парсинг) — доверяем первому аудитору
        return prior_verdict


# Правила обработки пунктов чек-листа.
# Ключ — индекс пункта (0 = пункт 1 в плане), значение — словарь:
#   file_keywords: список подстрок (lowercase) для поиска в имени файла; файл попадает в evidence если совпала любая.
#   include_template: добавлять ли текст самого "ПЛАН аудита" в evidence (по умолчанию True для всех пунктов в правилах).
#   extra_instructions: блок специфичных правил, дописывается в промпт verify_item_strict.
ITEM_RULES = {
    0: {
        "file_keywords": ["план"],
        "extra_instructions": (
            "Этот пункт проверяет САМ документ \"ПЛАН аудита\" (включён в evidence как файл \"ПЛАН АУДИТА (проверяемый документ)\").\n\n"
            "ПРОВЕРЬ:\n"
            "1. В шапке Плана АУДИТА должна быть дата утверждения/согласования (слова \"утверждаю\", \"утверждено\", \"согласовано\" и дата рядом).\n"
            "2. В Плане указаны сроки проведения аудита (даты начала и окончания).\n"
            "3. ПРАВИЛО: дата утверждения НЕ ПОЗДНЕЕ чем за 5 рабочих дней до даты начала аудита (СБ и ВС не считаются).\n\n"
            "NOK если: нет даты утверждения; нет сроков; интервал < 5 рабочих дней.\n"
            "В reason укажи обе даты и расчёт рабочих дней."
        ),
    },
    1: {
        "file_keywords": ["оргструктур", "орг.структур", "орг структур", "организационн"],
        "extra_instructions": (
            "Из файла \"ПЛАН аудита\" возьми указанную организационную структуру заявителя. "
            "Найди в источниках файлы по орг.структуре и сопоставь состав/иерархию. "
            "OK только если структуры совпадают по существу. NOK при расхождениях или отсутствии подтверждающего файла."
        ),
    },
    2: {
        "file_keywords": ["трудоемкост", "трудоёмкост"],
        "extra_instructions": (
            "Найди файл \"Трудоемкость\". Проверь данные о трудоёмкости (чел.-дни, распределение по площадкам/экспертам) "
            "и сопоставь с тем, что в Плане. NOK если данных нет или расходятся."
        ),
    },
    3: {
        "file_keywords": ["план", "приказ", "эг", "назначен"],
        "extra_instructions": (
            "В шаблоне Плана есть блок \"1 Цель и область аудита\" с галочками/крестиками — они переданы в evidence "
            "в секции \"ОТМЕТКИ\" / \"СИМВОЛЫ-ОТМЕТКИ\" / \"ЧЕКБОКСЫ PDF\". Найди файл \"Приказ о назначении ЭГ\" "
            "и проверь, что отметки в Плане соответствуют тому, что назначено в приказе. "
            "NOK если хотя бы одна отметка не соответствует или нужная не проставлена."
        ),
    },
    4: {
        "file_keywords": ["оквэд", "сертификат", "игс", "разбивк"],
        "extra_instructions": (
            "В шаблоне Плана найди строку НЕПОСРЕДСТВЕННО НАД \"область применения СМК (область сертификации)\". "
            "Эта строка должна ДОСЛОВНО (буква в букву) совпадать с формулировкой из файлов \"Разбивка кодов ОКВЭД\" "
            "и \"Сертификат ИГС\". NOK при ЛЮБЫХ расхождениях формулировки."
        ),
    },
    5: {
        "file_keywords": ["договор", "доп.соглашен", "доп соглашен", "доп. соглашен"],
        "extra_instructions": (
            "В шаблоне Плана пункт 4 — \"Наименование и адрес аудитируемых производственных площадок\". "
            "Дата и № договора (а также № и дата доп.соглашения, если есть) должны совпадать с файлом \"Договор\". "
            "NOK если номер/дата договора не совпали или не упомянуто доп.соглашение, когда оно реально присутствует в источниках."
        ),
    },
    6: {
        "file_keywords": ["заявка"],
        "extra_instructions": (
            "Сравни данные пункта чек-листа с информацией из файла \"Заявка\". "
            "NOK если данные расходятся или Заявки в источниках нет."
        ),
    },
    7: {
        "file_keywords": ["план", "приказ", "эг", "назначен"],
        "extra_instructions": (
            "В шаблоне Плана найди пункт 5 \"Сроки проведения аудита\" и сравни даты с файлом \"Приказ о назначении ЭГ\". "
            "NOK при любом расхождении дат."
        ),
    },
    8: {
        "file_keywords": ["план", "приказ", "эг", "назначен"],
        "extra_instructions": (
            "В шаблоне Плана найди раздел \"Состав экспертной группы\" (ФИО, роли). "
            "Сравни с составом из файла \"Приказ о назначении ЭГ\". NOK при расхождении состава или ролей."
        ),
    },
    9: {
        "file_keywords": ["акт", "предыдущ", "результат"],
        "extra_instructions": (
            "Найди файл с \"актом по результатам предыдущего аудита\" и сравни его данные с пунктом 8 шаблона Плана. "
            "NOK при расхождениях или отсутствии акта."
        ),
    },
    10: {
        "file_keywords": ["сто"],
        "extra_instructions": (
            "Список требуемых СТО берётся из самого Плана (упоминания СТО в чек-листе/тексте). "
            "Проверь, что в источниках есть отдельные файлы с \"СТО\" в имени для каждого упомянутого СТО. "
            "NOK если хотя бы один требуемый СТО отсутствует среди файлов."
        ),
    },
    11: {
        "file_keywords": ["план"],
        "extra_instructions": (
            "В шаблоне Плана проверь наличие совещаний:\n"
            "- \"Рабочее совещание\" и \"Промежуточное совещание\" в КОНЦЕ КАЖДОГО ДНЯ;\n"
            "- \"Предварительное совещание\" в НАЧАЛЕ КАЖДОГО ДНЯ на новой площадке (новый адрес);\n"
            "- \"Заключительное совещание\" в КОНЦЕ ПОСЛЕДНЕГО ДНЯ.\n"
            "NOK если хотя бы одно требуемое совещание не запланировано."
        ),
    },
    12: {
        "file_keywords": ["план"],
        "extra_instructions": (
            "В шаблоне Плана проверь наличие \"Инструктаж по технике безопасности и охране труда\" "
            "на КАЖДОЙ площадке и для КАЖДОГО эксперта. NOK если для какой-либо площадки или эксперта инструктаж пропущен."
        ),
    },
    13: {
        "file_keywords": ["акт", "предыдущ"],
        "extra_instructions": (
            "Сравни процессы из файла \"акт предыдущего аудита\" с колонкой \"Пункт стандарта\" в шаблоне Плана. "
            "Все процессы из акта должны быть охвачены в колонке Плана. NOK если хотя бы один процесс из акта отсутствует."
        ),
    },
    14: {
        "file_keywords": ["план"],
        "extra_instructions": (
            "В шаблоне Плана в колонке \"Пункт стандарта\": в каждой строке, где встречается слово \"Процесс\", "
            "должны быть указаны пункты 4.4.1, 4.4.2 и 4.4.3 (все три). "
            "NOK если в строке с \"Процесс\" отсутствует хотя бы один из 4.4.1 / 4.4.2 / 4.4.3."
        ),
    },
    15: {
        "file_keywords": ["план"],
        "extra_instructions": (
            "В шаблоне Плана найди временные интервалы перерыва (формат вида \"12:00 - 12:30\"). "
            "Каждый перерыв должен длиться 30 минут или больше (то есть >= 30 минут). "
            "Перерыв ровно 30 минут — это OK, требование выполнено. "
            "NOK ставь ТОЛЬКО если есть хотя бы один перерыв СТРОГО короче 30 минут (29 минут и меньше). "
            "В reason укажи найденные перерывы и их длительность в минутах."
        ),
    },
}


def _files_by_keyword(file_names: list, keywords: list) -> list:
    """Возвращает файлы, в имени которых встречается ЛЮБАЯ из подстрок (без учёта регистра)."""
    if not keywords:
        return []
    kws = [k.lower() for k in keywords]
    matched = []
    for f in file_names:
        lf = f.lower()
        if any(kw in lf for kw in kws):
            matched.append(f)
    return matched


def process_checklist_advanced(api_key: str, all_texts: dict,
                                checklist_structure: list,
                                ii_references: dict,
                                model: str = "GigaChat",
                                template_text: str = "") -> dict:
    """
    Per-item обработка чек-листа:
      1. Извлечение шапки (1 вызов)
      2. Для каждого пункта:
         a. Классификация релевантных файлов (1 вызов)
         b. Строгая проверка NOK-first (1 вызов)
         c. Adversarial-перепроверка если OK (1 вызов)
    Всего: 1 + N*(2..3) вызовов GigaChat.
    """
    global processing_status

    total = len(checklist_structure)
    file_names = list(all_texts.keys())

    processing_status.update({
        "stage": "header",
        "current": 0,
        "total": total,
        "message": "Извлечение шапки документа...",
        "detail": "",
    })

    try:
        header_data = extract_header_info(api_key, all_texts, model=model)
    except Exception as e:
        print(f"[header] Ошибка: {e}")
        header_data = {
            "Наименование Заявителя": "не найдено",
            "Вид аудита": "не найдено",
            "Даты проведения": "не найдено",
            "РЭГ": "не найдено",
        }

    checklist_results = []

    for idx, item in enumerate(checklist_structure):
        processing_status.update({
            "stage": "verify",
            "current": idx + 1,
            "message": f"Пункт {idx+1}/{total}: {item['area'][:60]}",
            "detail": "классификация файлов...",
        })

        try:
            rule = ITEM_RULES.get(idx)
            file_names = list(all_texts.keys())

            if rule and rule.get("file_keywords"):
                # Детерминированный отбор по ключевым подстрокам в имени файла
                relevant = _files_by_keyword(file_names, rule["file_keywords"])
                if not relevant:
                    # Фоллбэк: общий поиск по разделу/ключевым словам
                    relevant = find_relevant_files_for_item(item, all_texts)
            else:
                relevant = find_relevant_files_for_item(item, all_texts)

            print(f"[item {idx+1}/{total}] релевантных файлов: {len(relevant)}")
            evidence = build_evidence_pack(relevant, all_texts, max_chars=180000)

            extra_rules = ""
            if rule:
                extra_rules = rule.get("extra_instructions", "")
                # Для пунктов с правилами всегда добавляем текст самого Плана
                if template_text:
                    template_block = f"=== ФАЙЛ: ПЛАН АУДИТА (проверяемый документ) ===\n{template_text[:30000]}"
                    evidence = template_block + ("\n\n" + evidence if evidence else "")

            processing_status["detail"] = f"проверка по {len(relevant)} файлу(ам)..."
            verdict = verify_item_strict(api_key, item, ii_references, evidence, model=model,
                                          extra_instructions=extra_rules)

            if verdict.get("ok"):
                processing_status["detail"] = "adversarial-перепроверка OK-вердикта..."
                verdict = adversarial_recheck(api_key, item, ii_references, evidence, verdict, model=model)

            verdict["_files_checked"] = relevant
            # Дописываем в обоснование список файлов, в которых искалась информация
            if relevant:
                shown = relevant[:10]
                files_note = "; ".join(shown)
                if len(relevant) > len(shown):
                    files_note += f" (и ещё {len(relevant) - len(shown)})"
                suffix = f" Искалось в файлах ({len(relevant)}): {files_note}."
            else:
                suffix = " Искалось во всех загруженных файлах (релевантных по разделу/ключевым словам не найдено)."
            verdict["reason"] = (verdict.get("reason") or "").rstrip() + suffix
            checklist_results.append(verdict)

            print(f"[item {idx+1}/{total}] {'OK' if verdict['ok'] else 'NOK'}: {item['area'][:50]}")
        except Exception as e:
            print(f"[item {idx+1}] Ошибка анализа: {e}")
            checklist_results.append({
                "ok": False,
                "nok": True,
                "reason": f"Ошибка анализа: {str(e)[:120]}",
                "ii_data_found": ", ".join(item.get("ii_markers", [])) or "нет",
                "evidence_quote": "ошибка обработки",
                "source_file": "—",
                "_files_checked": [],
            })

    processing_status.update({
        "stage": "done",
        "current": total,
        "total": total,
        "message": "Обработка завершена",
        "detail": "",
    })

    return {"header": header_data, "checklist": checklist_results}


def validate_filled_document(output_path: str, checklist_structure: list[dict], extracted_data: dict) -> dict:
    """
    Проверка корректности заполненного документа.
    Возвращает отчёт о валидации.
    """
    doc = Document(output_path)
    issues = []
    warnings = []
    ok_count = 0
    nok_count = 0
    filled_count = 0
    total_checklist_rows = 0
    
    checklist_data = extracted_data.get("checklist", [])
    
    for table in doc.tables:
        first_row = table.rows[0]
        first_cell_text = first_row.cells[0].text.strip()
        cols_count = len(first_row.cells)
        
        # === Проверка шапки ===
        if cols_count == 2:
            header_data = extracted_data.get("header", {})
            for row in table.rows:
                cells_text = [cell.text.strip() for cell in row.cells]
                full_row_text = " ".join(cells_text)
                
                # Проверяем заполнено ли каждое поле шапки
                if any(key in full_row_text for key in ["Наименования Заявителя", "Вид аудита", "Даты проведения", "РЭГ"]):
                    # Есть ли данные в правой ячейке
                    right_cell = row.cells[1].text.strip() if len(row.cells) > 1 else ""
                    left_cell = row.cells[0].text.strip()
                    
                    # Определяем какое поле ожидаем
                    expected_field = None
                    if "Наименования Заявителя" in full_row_text:
                        expected_field = "Наименование Заявителя"
                    elif "Вид аудита" in full_row_text:
                        expected_field = "Вид аудита"
                    elif "Даты проведения" in full_row_text:
                        expected_field = "Даты проведения"
                    elif "РЭГ" in full_row_text and len(full_row_text) < 50:
                        expected_field = "РЭГ"
                    
                    if expected_field and expected_field in header_data:
                        # Проверяем что значение записано
                        combined = left_cell + " " + right_cell
                        if header_data[expected_field] not in combined and not right_cell:
                            warnings.append(f"Поле шапки '{expected_field}' возможно не заполнено")
        
        # === Проверка чек-листа ===
        if cols_count == 5 and "Область проверки" in first_cell_text:
            row_idx = 0
            for row in table.rows[1:]:  # Пропускаем заголовок
                row_idx += 1
                total_checklist_rows += 1
                
                area = row.cells[0].text.strip()
                ok_cell = row.cells[2].text.strip()
                nok_cell = row.cells[3].text.strip()
                prob_cell = row.cells[4].text.strip()
                
                if not area:
                    continue
                
                has_ok = "☑" in ok_cell or "OK" in ok_cell.upper()
                has_nok = "☒" in nok_cell or "NOK" in nok_cell.upper()
                
                # Проверяем что стоит ровно один маркер
                if has_ok and has_nok:
                    issues.append(f"Строка {row_idx} ('{area[:40]}...'): стоят ОБА маркера OK и NOK!")
                elif not has_ok and not has_nok:
                    issues.append(f"Строка {row_idx} ('{area[:40]}...'): НЕТ маркера OK или NOK!")
                elif has_ok:
                    ok_count += 1
                elif has_nok:
                    nok_count += 1
                
                filled_count += 1
    
    # Проверяем что все строки из структуры прошли валидацию
    if total_checklist_rows < len(checklist_structure):
        issues.append(f"В документе заполено {total_checklist_rows} строк, ожидалось {len(checklist_structure)}")
    
    # Проверяем количество пунктов от GigaChat
    if len(checklist_data) != len(checklist_structure):
        warnings.append(f"GigaChat вернул {len(checklist_data)} пунктов, в документе {len(checklist_structure)}")
    
    is_valid = len(issues) == 0
    
    return {
        "valid": is_valid,
        "issues": issues,
        "warnings": warnings,
        "stats": {
            "total_rows": total_checklist_rows,
            "expected_rows": len(checklist_structure),
            "ok_count": ok_count,
            "nok_count": nok_count,
            "filled_count": filled_count
        }
    }


@app.get("/")
async def root():
    return {"message": "Audit Plan Filler API", "version": "1.0.0"}


@app.get("/api/validate/{filename}")
async def validate_result(filename: str):
    """Проверка корректности заполненного документа"""
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Файл не найден")
    
    # Загружаем последнюю структуру чек-листа из сохранённых данных
    # Для этого прочитаем сохранённый результат если есть
    # Или просто вердим базовую валидацию
    try:
        doc = Document(str(file_path))
        checklist_rows = []
        header_filled = []
        
        for table in doc.tables:
            cols = len(table.rows[0].cells)
            
            if cols == 2:  # Шапка
                for row in table.rows:
                    text = " ".join(cell.text.strip() for cell in row.cells)
                    right = row.cells[1].text.strip() if len(row.cells) > 1 else ""
                    if any(k in text for k in ["Заявител", "Вид аудита", "Даты", "РЭГ"]):
                        header_filled.append(bool(right and len(right) > 2))
            
            if cols == 5:
                first_cell = table.rows[0].cells[0].text.strip()
                if "Область проверки" in first_cell:
                    for row in table.rows[1:]:
                        area = row.cells[0].text.strip()
                        ok = row.cells[2].text.strip()
                        nok = row.cells[3].text.strip()
                        if area:
                            has_ok = "☑" in ok
                            has_nok = "☒" in nok
                            checklist_rows.append({
                                "area": area[:60],
                                "ok": has_ok,
                                "nok": has_nok,
                                "both": has_ok and has_nok,
                                "none": not has_ok and not has_nok
                            })
        
        issues = []
        for i, row in enumerate(checklist_rows):
            if row["both"]:
                issues.append(f"Строка {i+1}: ОБА маркера OK и NOK")
            elif row["none"]:
                issues.append(f"Строка {i+1}: НЕТ маркера")
        
        return {
            "valid": len(issues) == 0,
            "header_filled": f"{sum(header_filled)}/{len(header_filled)} полей",
            "checklist_total": len(checklist_rows),
            "ok_count": sum(1 for r in checklist_rows if r["ok"]),
            "nok_count": sum(1 for r in checklist_rows if r["nok"]),
            "issues": issues
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка валидации: {str(e)}")


@app.post("/api/settings/gigachat")
async def save_gigachat_settings(settings: GigaChatSettings):
    """Сохранение настроек GigaChat (ключ + выбранная модель)"""
    settings_file = Path(__file__).parent / "gigachat_settings.json"
    existing = _load_settings()
    existing["api_key"] = settings.api_key
    if settings.model is not None:
        existing["model"] = settings.model.strip()
    with open(settings_file, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    return {"status": "ok", "message": "Настройки сохранены"}


@app.get("/api/settings/gigachat")
async def get_gigachat_settings():
    """Получение настроек GigaChat"""
    settings = _load_settings()
    return {
        "api_key": settings.get("api_key", ""),
        "model": settings.get("model", "") or DEFAULT_GIGACHAT_MODEL,
        "default_model": DEFAULT_GIGACHAT_MODEL,
    }


@app.post("/api/session/new")
async def create_session():
    """Создать новую пустую сессию загрузок."""
    sid, _ = _resolve_session_dir(None)
    return {"status": "ok", "session_id": sid}


@app.post("/api/upload")
async def upload_files(
    files: list[UploadFile] = File(...),
    session_id: Optional[str] = Form(None),
):
    """Загрузка документов (docx, docm, pdf, xlsx) в папку сессии."""
    sid, sdir = _resolve_session_dir(session_id)

    uploaded = []
    for file in files:
        if not file.filename:
            continue

        file_path = sdir / file.filename
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        uploaded.append({
            "filename": file.filename,
            "path": str(file_path),
            "size": file_path.stat().st_size
        })

    return {"status": "ok", "session_id": sid, "uploaded_files": uploaded, "count": len(uploaded)}


@app.post("/api/process")
async def process_documents(
    api_key: str = Form(...),
    template_file: str = Form(...),
    session_id: str = Form(...),
    plan_doc_file: Optional[str] = Form(None),
):
    """
    Обработка документов через GigaChat и заполнение плана
    """
    if not _SESSION_ID_RE.match(session_id or ""):
        raise HTTPException(status_code=400, detail="Некорректный session_id")
    sdir = UPLOAD_DIR / session_id
    if not sdir.is_dir():
        raise HTTPException(status_code=400, detail="Сессия не найдена")

    # Проверяем наличие загруженных файлов в папке сессии
    docx_files = list(sdir.glob("*.docx"))
    docm_files = list(sdir.glob("*.docm"))
    pdf_files = list(sdir.glob("*.pdf"))
    xlsx_files = list(sdir.glob("*.xlsx"))
    all_docs = docx_files + docm_files + pdf_files + xlsx_files

    if not all_docs:
        raise HTTPException(status_code=400, detail="Нет загруженных документов")

    # Находим шаблон плана
    template_path = sdir / template_file
    if not template_path.exists():
        # Ищем среди стандартных шаблонов
        possible_names = ["ИИ шаблон плана.docm", "ИИ -ЧК -План АУДИТА.docx"]
        for name in possible_names:
            alt_path = sdir / name
            if alt_path.exists():
                template_path = alt_path
                break
        else:
            raise HTTPException(status_code=400, detail=f"Шаблон '{template_file}' не найден")

    # Извлекаем текст из всех документов КРОМЕ шаблона.
    # ВАЖНО: эта работа CPU-bound (особенно OCR) — выносим в thread, иначе блокируется
    # event loop и /api/status перестаёт отвечать (фронт не видит прогресс).
    import asyncio

    source_docs = [d for d in all_docs if d.name != template_file]
    total_files = len(source_docs)
    processing_status.update({
        "stage": "extract",
        "current": 0,
        "total": total_files,
        "message": f"Извлечение текста из {total_files} файлов",
        "detail": "",
    })

    EXTRACT_WORKERS = 3

    def _extract_one(doc_file: Path):
        """Извлечь текст одного файла. Возвращает (filename, text) или (filename, None) при ошибке."""
        suffix = doc_file.suffix.lower()
        try:
            if suffix in ('.docx', '.docm'):
                text = extract_text_from_docx(str(doc_file))
            elif suffix == '.pdf':
                text = extract_text_from_pdf(str(doc_file))
            elif suffix == '.xlsx':
                text = extract_text_from_xlsx(str(doc_file))
            else:
                print(f"Пропущен файл {doc_file.name} — неподдерживаемый формат")
                return (doc_file.name, None)
            return (doc_file.name, text)
        except Exception as e:
            print(f"ERR Ошибка чтения {doc_file.name}: {e}")
            return (doc_file.name, None)

    def _extract_all_texts_parallel():
        """Параллельное извлечение в пуле потоков (3 файла одновременно)."""
        from concurrent.futures import ThreadPoolExecutor, as_completed
        import threading

        all_texts = {}
        analyzed_files = []
        in_progress = set()
        lock = threading.Lock()
        completed = 0

        def update_status():
            with lock:
                names = ", ".join(sorted(in_progress)[:3])
                processing_status.update({
                    "stage": "extract",
                    "current": completed,
                    "total": total_files,
                    "message": f"Файлы {completed}/{total_files} (параллельно x{EXTRACT_WORKERS})",
                    "detail": f"в обработке: {names}" if names else "",
                })

        with ThreadPoolExecutor(max_workers=EXTRACT_WORKERS) as ex:
            future_to_file = {}
            for doc_file in source_docs:
                fut = ex.submit(_extract_one, doc_file)
                future_to_file[fut] = doc_file
                with lock:
                    in_progress.add(doc_file.name[:60])

            update_status()

            for fut in as_completed(future_to_file):
                doc_file = future_to_file[fut]
                fname, text = fut.result()
                with lock:
                    in_progress.discard(doc_file.name[:60])
                    completed += 1
                    if text is not None:
                        all_texts[fname] = text
                        analyzed_files.append(fname)
                        print(f"OK Проанализирован ({completed}/{total_files}): {fname}")
                update_status()

        return all_texts, analyzed_files

    all_texts, analyzed_files = await asyncio.to_thread(_extract_all_texts_parallel)

    print(f"\nВсего проанализировано файлов: {len(analyzed_files)}")
    print(f"Файлы: {analyzed_files}\n")

    processing_status.update({
        "stage": "extract",
        "current": total_files,
        "total": total_files,
        "message": "Чтение шаблона чек-листа...",
        "detail": "",
    })

    # Текст самого шаблона (нужен для проверки пункта 1 — даты в шапке Плана)
    try:
        template_text = await asyncio.to_thread(extract_text_from_docx, str(template_path))
    except Exception as e:
        print(f"WARN: Не удалось извлечь текст шаблона для пункта 1: {e}")
        template_text = ""

    # Опциональный отдельный файл "План"
    plan_doc_text = ""
    if plan_doc_file:
        plan_doc_path = sdir / plan_doc_file
        if not plan_doc_path.exists():
            raise HTTPException(status_code=400, detail=f"Файл 'План' '{plan_doc_file}' не найден в сессии")
        processing_status.update({
            "stage": "extract",
            "message": f"Чтение файла 'План': {plan_doc_file[:80]}",
            "detail": "",
        })

        def _read_plan_doc():
            suffix = plan_doc_path.suffix.lower()
            if suffix in ('.docx', '.docm'):
                return extract_text_from_docx(str(plan_doc_path))
            elif suffix == '.pdf':
                return extract_text_from_pdf(str(plan_doc_path))
            elif suffix == '.xlsx':
                return extract_text_from_xlsx(str(plan_doc_path))
            return ""

        try:
            plan_doc_text = await asyncio.to_thread(_read_plan_doc)
        except Exception as e:
            print(f"WARN: Не удалось прочитать файл 'План' {plan_doc_file}: {e}")
            plan_doc_text = ""

    # Что использовать как "текст Плана" для evidence пунктов с правилами:
    # приоритет — отдельный plan_doc_text, иначе template_text.
    plan_text_for_items = plan_doc_text or template_text

    # Извлекаем структуру чек-листа из шаблона
    checklist_structure = extract_checklist_from_template(str(template_path))
    doc_checklist_count = len(checklist_structure)

    # Извлекаем ВСЕ маркеры ИИ из шаблона
    ii_references = extract_ii_references(str(template_path))
    print(f"Найдено маркеров ИИ в шаблоне: {list(ii_references.keys())}")
    print(f"Пунктов чек-листа: {doc_checklist_count}")

    # Активная модель из настроек (с фолбэком на дефолт)
    active_model = _get_active_model()

    # Pre-flight: быстрая проверка GigaChat. Если auth/связь упали — ошибаемся сразу,
    # не гоняя 16 пунктов впустую.
    processing_status.update({
        "stage": "preflight",
        "current": 0,
        "total": 1,
        "message": f"Проверка соединения с GigaChat ({active_model})...",
        "detail": "",
    })
    print("Preflight GigaChat...")
    preflight = await asyncio.to_thread(gigachat_preflight, api_key, 8.0, active_model)
    if not preflight["ok"]:
        raise HTTPException(
            status_code=503,
            detail=f"GigaChat недоступен ({preflight['stage']}): {preflight['detail']}"
        )
    print(f"Preflight OK. Доступные модели: {preflight['models']}")
    print(f"Запуск per-item пайплайна (модель: {active_model})...")

    # НОВЫЙ ПАЙПЛАЙН: per-item проверка с NOK-first + adversarial
    try:
        extracted_data = await asyncio.to_thread(
            process_checklist_advanced,
            api_key, all_texts, checklist_structure, ii_references, active_model,
            plan_text_for_items
        )
    except Exception as e:
        processing_status.update({"stage": "error", "message": str(e)[:200]})
        return ProcessingResult(status="error", message=f"Ошибка пайплайна: {str(e)}")

    if "error" in extracted_data:
        return ProcessingResult(status="error", message=extracted_data["error"])
    
    # Приводим количество пунктов чек-листа к реальному в документе
    if "checklist" in extracted_data:
        actual = extracted_data["checklist"]
        if len(actual) > doc_checklist_count:
            # GigaChat вернул больше чем строк в документе — обрезаем
            print(f"Предупреждение: GigaChat вернул {len(actual)} пунктов, в документе {doc_checklist_count}. Обрезано.")
            extracted_data["checklist"] = actual[:doc_checklist_count]
        elif len(actual) < doc_checklist_count:
            # GigaChat вернул меньше — дополняем заглушками
            for i in range(len(actual), doc_checklist_count):
                area_name = checklist_structure[i]['area'][:60] if i < len(checklist_structure) else f"пункт {i+1}"
                ii_markers = checklist_structure[i].get('ii_markers', [])
                extracted_data["checklist"].append({
                    "ok": False,
                    "nok": True,
                    "reason": f"Нет данных для проверки: {area_name}",
                    "problems": f"Нет данных для проверки: {area_name}",
                    "ii_data_found": ', '.join(ii_markers) if ii_markers else "нет маркеров ИИ"
                })
        else:
            # Убедимся что у всех есть поля reason и ii_data_found
            for item in extracted_data["checklist"]:
                if "reason" not in item:
                    item["reason"] = item.get("problems", "")
                if "ii_data_found" not in item:
                    item["ii_data_found"] = ""

    # Формируем выходной файл
    output_filename = f"Заполненный_План_АУДИТА.docx"
    output_path = OUTPUT_DIR / output_filename

    # Заполняем шаблон (шапка + чек-лист OK/NOK + Проблемные зоны)
    processing_status.update({
        "stage": "fill",
        "message": "Заполнение шаблона результатами...",
        "detail": "",
    })
    try:
        await asyncio.to_thread(
            fill_plan_with_checklist,
            str(template_path), extracted_data, str(output_path), checklist_structure
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка заполнения плана: {str(e)}")
    
    # === ВАЛИДАЦИЯ: проверяем 100% корректность ===
    processing_status.update({
        "stage": "validate",
        "message": "Валидация результата...",
        "detail": "",
    })
    validation = await validate_result("Заполненный_План_АУДИТА.docx")
    
    # Если есть критические проблемы — пробуем исправить
    if not validation.get("valid"):
        print(f"Валидация: обнаружены проблемы {validation.get('issues')}")
        # В будущем здесь можно добавить авто-исправление
        # Пока просто возвращаем отчёт
    
    return ProcessingResult(
        status="success",
        message=f"План заполнен. Проанализировано файлов: {len(analyzed_files)}",
        extracted_data=extracted_data,
        output_file=output_filename,
        validation=validation,
        analyzed_files=analyzed_files
    )


@app.get("/api/download/{filename}")
async def download_file(filename: str):
    """Скачивание готового файла"""
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Файл не найден")
    
    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.get("/api/files")
async def list_files(session_id: Optional[str] = None):
    """Список загруженных файлов в сессии."""
    if not session_id or not _SESSION_ID_RE.match(session_id):
        raise HTTPException(status_code=400, detail="Некорректный session_id")
    sdir = UPLOAD_DIR / session_id
    if not sdir.is_dir():
        return {"files": []}
    files = []
    for f in sdir.iterdir():
        if f.is_file():
            files.append({
                "name": f.name,
                "size": f.stat().st_size,
                "type": "template" if "шаблон" in f.name.lower() or "план" in f.name.lower() else "source"
            })
    return {"files": files}


@app.delete("/api/files/{filename}")
async def delete_file(filename: str, session_id: Optional[str] = None):
    """Удаление загруженного файла из папки сессии."""
    if not session_id or not _SESSION_ID_RE.match(session_id):
        raise HTTPException(status_code=400, detail="Некорректный session_id")
    file_path = UPLOAD_DIR / session_id / filename
    if file_path.exists():
        file_path.unlink()
        return {"status": "ok", "message": f"Файл {filename} удален"}
    raise HTTPException(status_code=404, detail="Файл не найден")


@app.post("/api/upload-from-path")
async def upload_from_path(
    source_path: str = Form(...),
    session_id: Optional[str] = Form(None),
):
    """Загрузка документов из указанной папки (все форматы) в папку сессии."""
    if os.environ.get("ENABLE_LOCAL_UPLOAD", "0") != "1":
        raise HTTPException(status_code=404, detail="Endpoint disabled")
    path = Path(source_path)
    if not path.exists():
        raise HTTPException(status_code=400, detail="Путь не существует")

    sid, sdir = _resolve_session_dir(session_id)

    uploaded = []
    extensions = ['*.docx', '*.docm', '*.pdf', '*.xlsx']
    for ext in extensions:
        for doc_file in path.glob(ext):
            dest = sdir / doc_file.name
            shutil.copy2(str(doc_file), str(dest))
            uploaded.append({
                "filename": doc_file.name,
                "path": str(dest),
                "size": dest.stat().st_size
            })

    return {"status": "ok", "session_id": sid, "uploaded_files": uploaded, "count": len(uploaded)}


@app.get("/api/status")
def get_processing_status():
    """Текущий статус обработки (для polling из UI во время /api/process)."""
    return processing_status


@app.get("/api/gigachat/diagnose")
def diagnose_gigachat():
    """
    Диагностика связи с GigaChat:
      - OAuth на 9443
      - Список доступных моделей
      - Проверка что DEFAULT_GIGACHAT_MODEL есть в списке
    Читает ключ из сохранённых настроек.
    """
    settings = _load_settings()
    key = settings.get("api_key", "")
    active_model = _get_active_model()
    if not key:
        return {"ok": False, "stage": "no_key",
                "detail": "API ключ не сохранён. Сначала сохраните ключ в настройках.",
                "models": [], "current_model": active_model}
    return gigachat_preflight(key, timeout=8.0, model=active_model)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
