"""
Извлечение полей документов пакета аудита по словарю кодов (docs/ЧЕК-ЛИСТЫ.md, §3).

На вход — текст из main.extract_text_from_*: строки таблиц docx склеены через « | »,
колонтитулы и отметки (галочки) вынесены в служебные секции «=== … ===». Сканы приходят
через OCR. Поле, которое не нашлось, возвращается как None — что делать с «нет данных»,
решает проверка (formula_checks.py), а не извлекатель.
"""

import re
from datetime import date
from typing import Optional

MONTHS_RU = {
    "январ": 1, "феврал": 2, "март": 3, "апрел": 4, "ма": 5, "июн": 6,
    "июл": 7, "август": 8, "сентябр": 9, "октябр": 10, "ноябр": 11, "декабр": 12,
}


def ru_month(word: str) -> Optional[int]:
    w = (word or "").lower().replace("ё", "е")
    for stem, n in MONTHS_RU.items():
        if w.startswith(stem):
            return n
    return None


def _mk_date(y: int, m: int, d: int) -> Optional[date]:
    if y < 100:
        y += 2000
    try:
        return date(y, m, d)
    except ValueError:
        return None


_NUM_DATE_RE = re.compile(r"(?<!\d)(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})(?!\d)")
_WORD_DATE_RE = re.compile(r"«?\s*(\d{1,2})\s*»?\s+([А-Яа-яЁё]{3,})\s+(\d{4})")


def find_dates(text: str) -> list:
    """Все даты в тексте в порядке появления: 30.07.2026, «12» августа 2026 г."""
    found = []
    for m in _NUM_DATE_RE.finditer(text or ""):
        d = _mk_date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        if d:
            found.append((m.start(), d))
    for m in _WORD_DATE_RE.finditer(text or ""):
        mon = ru_month(m.group(2))
        d = _mk_date(int(m.group(3)), mon, int(m.group(1))) if mon else None
        if d:
            found.append((m.start(), d))
    return [d for _, d in sorted(found, key=lambda x: x[0])]


def first_date(text: str) -> Optional[date]:
    dates = find_dates(text)
    return dates[0] if dates else None


def find_period(text: str) -> Optional[tuple]:
    """Период аудита (начало, конец): «27-30.07.2026», «27.07.2026 - 30.07.2026»,
    «с 27.07.2026 по 30.07.2026». Одна дата — период из одного дня."""
    t = text or ""
    m = re.search(r"(\d{1,2})\s*[-–—]\s*(\d{1,2})\.(\d{1,2})\.(\d{2,4})", t)
    if m:
        y, mo = int(m.group(4)), int(m.group(3))
        a, b = _mk_date(y, mo, int(m.group(1))), _mk_date(y, mo, int(m.group(2)))
        if a and b:
            return (a, b)
    m = re.search(r"(\d{1,2}\.\d{1,2}\.\d{2,4})\s*(?:[-–—]|по)\s*(\d{1,2}\.\d{1,2}\.\d{2,4})", t)
    if m:
        a, b = first_date(m.group(1)), first_date(m.group(2))
        if a and b:
            return (a, b)
    d = first_date(t)
    return (d, d) if d else None


def fmt_date(d) -> str:
    return d.strftime("%d.%m.%Y") if d else "—"


def fmt_period(p) -> str:
    if not p:
        return "—"
    return fmt_date(p[0]) if p[0] == p[1] else f"{fmt_date(p[0])}–{fmt_date(p[1])}"


def workday_deadline(signed, n: int):
    """Срок «n рабочих дней с даты подписания»: день подписания — первый рабочий день
    (если он рабочий), так считает ОС (решение методолога, 12.10.2026).
    Акт от 30.07.2026 (чт): 5 р.д. — до 05.08.2026, 20 р.д. — до 26.08.2026."""
    return add_workdays(signed, n - 1) if signed.weekday() < 5 else add_workdays(signed, n)


def add_workdays(d, n: int):
    """d ± n рабочих дней (Пн–Пт)."""
    from datetime import timedelta
    step = 1 if n >= 0 else -1
    left, cur = abs(n), d
    while left:
        cur += timedelta(days=step)
        if cur.weekday() < 5:
            left -= 1
    return cur


# ---------- Наименование организации ----------

OPF_EXPANSIONS = [
    ("публичное акционерное общество", "пао"),
    ("непубличное акционерное общество", "нао"),
    ("акционерное общество", "ао"),
    ("общество с ограниченной ответственностью", "ооо"),
    ("закрытое акционерное общество", "зао"),
    ("открытое акционерное общество", "оао"),
    ("индивидуальный предприниматель", "ип"),
]


def normalize_org_name(name: str) -> tuple[str, str]:
    """Нормализует наименование юр.лица.

    Возвращает (opf, core) — токен ОПФ (ооо/ао/пао/...) и «ядро» (название без ОПФ, без кавычек/пробелов).
    Любое отклонение в ОПФ (ПААО vs ПАО) или в ядре (Ромашкаа vs Ромашка) даст разные значения.
    """
    if not name:
        return ("", "")
    s = name.lower().strip()
    # убираем все виды кавычек и лишние знаки
    for ch in ['«', '»', '"', '"', '"', "'", '`', '(', ')', '\n', '\t']:
        s = s.replace(ch, ' ')
    # схлопываем пробелы
    s = re.sub(r'\s+', ' ', s).strip()

    # развёрнутая ОПФ → аббревиатура (сравниваем полные формы до коротких, чтобы «акционерное общество» не съело «публичное акционерное общество»)
    opf = ""
    for full, short in OPF_EXPANSIONS:
        if s.startswith(full + ' ') or s == full:
            opf = short
            s = s[len(full):].strip()
            break
    if not opf:
        # Если ОПФ-аббревиатура — первое короткое «слово» (2–6 букв) перед названием.
        # Сюда попадает и валидное «ПАО», и опечатки вида «ПААО» — то, что и нужно ловить.
        parts = s.split(' ', 1)
        head = parts[0]
        if 2 <= len(head) <= 6 and head.isalpha():
            opf = head
            s = parts[1] if len(parts) > 1 else ''

    # ядро — буквы и цифры (выкидываем пунктуацию и пробелы для устойчивого сравнения)
    core = re.sub(r'[^0-9a-zа-яё]+', '', s)
    return (opf, core)


_OPF_ALT = "|".join(
    [re.escape(full) for full, _ in OPF_EXPANSIONS]
    + [r"ООО", r"ПАО", r"НАО", r"ЗАО", r"ОАО", r"АО", r"ИП"]
)
# ОПФ и открывающая кавычка; само название разбирает _quoted_name
_ORG_RE = re.compile(r"(?<![А-Яа-яЁё])(" + _OPF_ALT + r")\s*(?=[«\"])", re.IGNORECASE)

# Орган по сертификации упоминается почти в каждом документе — это не проверяемая организация
_CERT_BODY_CORE = "кцперспектива"


def _quoted_name(text: str, start: int, limit: int = 150) -> Optional[str]:
    """Название в кавычках с позиции открывающей кавычки, с учётом вложенности:
    «КЦ «Перспектива»». В документах встречается и незакрытая внешняя кавычка
    («КЦ «Перспектива» 1. Область…») — тогда название кончается на последней
    закрывающей кавычке до перевода строки или limit символов, а не на следующей
    кавычке через страницу текста."""
    if text[start] == '"':
        end = text.find('"', start + 1, start + limit)
        return text[start + 1: end] if end > 0 and "\n" not in text[start: end] else None
    depth, nested_close = 0, None
    for i in range(start, min(len(text), start + limit)):
        ch = text[i]
        if ch == "\n":
            break
        if ch == "«":
            depth += 1
        elif ch == "»":
            depth -= 1
            if depth == 0:
                return text[start + 1: i]
            if depth == 1 and nested_close is None:
                nested_close = i  # закрылась вложенная кавычка: «КЦ «Перспектива»
    # внешняя кавычка не закрыта — название до первой закрытой вложенной
    return text[start + 1: nested_close + 1] if nested_close else None


def find_org(text: str, last: bool = False) -> Optional[dict]:
    """Организация в тексте (первая, при last=True — последняя), кроме органа по
    сертификации: {"name": как написано, "key": (opf, core), "full_form": ОПФ полностью}.
    last=True — для блока над подписью «наименование проверяемой организации»: выше него
    может стоять «УТВЕРЖДАЮ … ООО «…»», а нужное поле — ближайшее к подписи."""
    t = text or ""
    orgs = []
    for m in _ORG_RE.finditer(t):
        inner = _quoted_name(t, m.end())
        if not inner or not inner.strip():
            continue
        name = f"{m.group(1)} «{inner.strip()}»"
        key = normalize_org_name(name)
        if key[1] == _CERT_BODY_CORE:
            continue
        orgs.append({"name": re.sub(r"\s+", " ", name), "key": key, "full_form": len(m.group(1)) > 5})
    if not orgs:
        return None
    return orgs[-1] if last else orgs[0]


# ---------- ФИО ----------

_SURNAME = r"([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)(?![\dА-Яа-яЁё])"
_PERSON_RES = [
    # Фамилия И.О.
    re.compile(r"(?<![А-Яа-яЁё])" + _SURNAME + r"\s+([А-ЯЁ])\.\s*([А-ЯЁ])\.?"),
    # И.О. Фамилия
    re.compile(r"(?<![А-Яа-яЁё])([А-ЯЁ])\.\s*([А-ЯЁ])\.\s*" + _SURNAME),
]


def find_persons(text: str) -> list:
    """ФИО в порядке появления как ключи «фамилия ио»: «Сукорская О.А.» и
    «О.А. Сукорская» дают одно и то же «сукорская оа». Форма «И.О. Фамилия» не
    ищется внутри уже найденной «Фамилия И.О.»: в «Сукорская О.А. Гревцов К.В.»
    нет человека «О.А. Гревцов»."""
    found = []
    t = text or ""
    taken = []
    for m in _PERSON_RES[0].finditer(t):
        found.append((m.start(), f"{m.group(1).lower()} {m.group(2).lower()}{m.group(3).lower()}"))
        taken.append(m.span())
    for m in _PERSON_RES[1].finditer(t):
        if any(a <= m.start() < b for a, b in taken):
            continue
        found.append((m.start(), f"{m.group(3).lower()} {m.group(1).lower()}{m.group(2).lower()}"))
    found.sort()
    keys = []
    for _, k in found:
        if k not in keys:
            keys.append(k)
    return keys


def person_label(key: str) -> str:
    """«сукорская оа» → «Сукорская О.А.»"""
    if not key:
        return "—"
    surname, _, ini = key.partition(" ")
    return surname.capitalize() + " " + "".join(f"{c.upper()}." for c in ini)


# ---------- Отметки ----------

def checkbox_marks(text: str) -> list:
    """Отметки из секции экстрактора «=== ОТМЕТКИ … ===»: [(отмечено, подпись)]."""
    i = (text or "").find("=== ОТМЕТКИ")
    if i < 0:
        return []
    marks = []
    for line in text[i:].splitlines()[1:]:
        if line.startswith("==="):
            break
        m = re.match(r"\s*([☐☑☒])\s*(?:НЕ\s+)?ОТМЕЧЕНО\s*\|\s*(.*)", line)
        if m:
            marks.append((m.group(1) != "☐", m.group(2).strip()))
    return marks


def _segment(text: str, start_pat: str, end_pat: Optional[str] = None, limit: int = 600,
             next_line: bool = False) -> str:
    """Кусок текста от start_pat (регулярка, без учёта регистра) до end_pat или limit символов.
    next_line — начать со следующей строки (пропустить хвост строки-заголовка раздела)."""
    m = re.search(start_pat, text or "", re.IGNORECASE)
    if not m:
        return ""
    start = m.end()
    if next_line:
        nl = text.find("\n", start)
        start = nl + 1 if nl >= 0 else len(text)
    seg = text[start: start + limit]
    if end_pat:
        e = re.search(end_pat, seg, re.IGNORECASE)
        if e:
            seg = seg[: e.start()]
    return seg


# ---------- Состав экспертной группы ----------

_ROLE_ORDER = [  # порядок важен: «Руководитель экспертной группы» содержит «эксперт»
    ("leader", "руководител"),
    ("tech", "техническ"),
    ("candidates", "кандидат"),
    ("observers", "наблюдател"),
    ("trainees", "стажер"),
    ("experts", "эксперт"),
]

ROLE_TITLES = {
    "leader": "руководитель ЭГ", "experts": "эксперт", "tech": "технический эксперт",
    "candidates": "кандидат в эксперты", "observers": "наблюдатель", "trainees": "стажёр",
}


def parse_team(segment: str) -> dict:
    """Состав ЭГ по строкам «8.1 Руководитель экспертной группы | Сукорская О.А.»:
    {роль: [ключи ФИО]}."""
    team = {}
    for line in (segment or "").splitlines():
        head, _, tail = line.partition("|")
        low = head.lower().replace("ё", "е")
        role = next((r for r, k in _ROLE_ORDER if k in low), None)
        if role is None:
            continue
        team.setdefault(role, []).extend(find_persons(tail))
    return team


# ---------- Общие разборщики разделов ----------

# Вид аудита: (канон, подстроки). «Расширение области» — сопровождение основного вида.
AUDIT_TYPES = [
    ("первый этап", ("первый этап", "первого этапа", "1 этап", "1-й этап")),
    ("второй этап", ("второй этап", "второго этапа", "2 этап", "2-й этап")),
    ("внеплановый ИК", ("внепланов",)),
    ("первый ИК", ("первый инспекцион", "первого инспекцион", "1 ик", "1ик")),
    ("второй ИК", ("второй инспекцион", "второго инспекцион", "2 ик", "2ик")),
    ("ресертификация", ("ресертификац",)),
    ("дополнительный", ("дополнительный аудит", "дополнительного аудита")),
    ("расширение области", ("расширение области", "расширения области")),
]
INSPECTION_TYPES = {"первый ИК", "второй ИК", "внеплановый ИК"}
EXTENSION = "расширение области"


def audit_types_in(label: str) -> set:
    low = (label or "").lower().replace("ё", "е")
    return {canon for canon, keys in AUDIT_TYPES if any(k in low for k in keys)}


def marked_audit_types(text: str) -> set:
    """Виды аудита, отмеченные галочкой (секция ОТМЕТКИ экстрактора)."""
    found = set()
    for checked, label in checkbox_marks(text):
        if checked:
            found |= audit_types_in(label)
    return found


def main_audit_type(text: str) -> Optional[str]:
    """Вид аудита, названный в тексте (Приказ ЭГ: «Для проведения второго инспекционного
    контроля»). Расширение области — не вид аудита, а сопровождение; ресертификация
    приоритетнее «первого этапа» («первого этапа ресертификационного аудита»)."""
    found = audit_types_in(text) - {EXTENSION}
    for canon in ("ресертификация", "второй ИК", "первый ИК", "внеплановый ИК",
                  "второй этап", "первый этап", "дополнительный"):
        if canon in found:
            return canon
    return None


def is_marked(text: str, phrase: str, exact: bool = False) -> Optional[bool]:
    """Отмечена ли галочка с подписью phrase (None — такой галочки нет). exact — подпись
    начинается с phrase и дальше нет значимых слов («Выдать сертификат» ≠ «Выдать
    сертификат после устранения…»): допускаются только коды разметки образцов."""
    ph = phrase.lower()
    result = None
    for checked, label in checkbox_marks(text):
        low = label.lower().replace("ё", "е")
        if exact:
            rest = low[len(ph):].strip() if low.startswith(ph) else None
            if rest is None or re.search(r"[а-я]{3,}", rest):
                continue
        elif ph not in low:
            continue
        result = bool(result) or checked
    return result


def exclusion_block(text: str) -> Optional[str]:
    """Раздел «Исключения из требований стандарта(ов)» — до следующего нумерованного
    раздела («9 …», «10 …») или 700 символов. Раньше брали 2000 символов — это захватывало
    соседний раздел графика с пунктами процессов (4.4.1, 4.3.2) и давало ложный NOK."""
    m = re.search(r"исключения из требований", text or "", re.IGNORECASE)
    if not m:
        return None
    raw_block = text[m.start(): m.start() + 2000]
    next_section = re.search(r'\n\s*(?:9|10|11|12)\s+[А-ЯA-Z]', raw_block)
    return raw_block[: next_section.start()] if next_section else raw_block[:700]


def exclusion_clauses(block: str) -> list:
    """Пункты СТО в разделе исключений — только те, рядом с которыми есть контекст
    исключения ('п.', 'пункт', 'абзац', 'не применяется', 'исключ'). Без дублей."""
    found = []
    for p in re.findall(r'\b(\d{1,2}(?:\.\d{1,2}){2,4})\b', block or ""):
        pos = block.find(p)
        ctx = block[max(0, pos - 40): pos + 80].lower()
        if any(k in ctx for k in ("п.", "пункт", "абзац", "не применя", "исключ")):
            found.append(p)
    return list(dict.fromkeys(found))


_STOP_WORDS = {"основная", "площадка", "российская", "федерация", "улица", "город", "область",
               "дом", "строение", "корпус", "офис", "район", "край"}


def word_set(text: str, min_len: int = 4) -> set:
    """Значимые слова (кириллица/латиница/цифры) длиной ≥ min_len."""
    s = re.sub(r"[^а-яa-z0-9 ]", " ", (text or "").lower().replace("ё", "е"))
    return {w for w in s.split() if len(w) >= min_len}


def jaccard(a: set, b: set) -> float:
    return len(a & b) / max(1, len(a | b))


def basis_refs(segment: str) -> tuple:
    """Реквизиты основания (договор, доп. соглашения, заявка): (номера, даты)."""
    nums = {re.sub(r"[^0-9a-zа-я/]", "", n.lower()).strip("/")
            for n in re.findall(r"№\s*([A-Za-zА-Яа-я0-9/\-_.]+)", segment or "")}
    return ({n for n in nums if n}, set(find_dates(segment)))


def address_tokens(segment: str) -> tuple:
    """(почтовые индексы, значимые слова адреса без служебных: «площадка», «ул.» …)."""
    postals = set(re.findall(r"(?<!\d)\d{6}(?!\d)", segment or ""))
    words = {w for w in word_set(segment, 3) if w not in _STOP_WORDS and not re.fullmatch(r"\d{6}", w)}
    return postals, words


def okved_codes(segment: str) -> list:
    return list(dict.fromkeys(re.findall(r"(?<![\d.])\d{2}\.\d{2}(?:\.\d{1,2})*(?![\d])", segment or "")))


def okved_group(code: str) -> str:
    """26.51.7 → 26.51: Акт и Сводный акт указывают коды на уровне группировки."""
    return ".".join(code.split(".")[:2])


# ---------- Акт по результатам аудита (А, Акт) ----------

def akt_fields(text: str) -> dict:
    """Поля Акта по результатам аудита.

    org (А20 / Акт1), types (Акт7), area (А30), basis (А31), sites (А32), okved (А33),
    exclusions (А34), compliance (А35), period (А8), team (А21), representative (А23),
    date (Акт2), leader_sign (Акт3), nc_23_2 / nc_total (Акт5), nc_significant / nc_minor
    (А36 / А37), pkd_deadline (Акт-ПКД1, п.17), kd_deadline (Акт6, п.18), extend / narrow
    (А40 / А41), distribution_org (Акт9), signers (подписи членов ЭГ).
    """
    t = text or ""
    body = t.split("=== ОТМЕТКИ")[0]
    f = {}
    org_block = _segment(t, r"\(вид системы менеджмента\)", r"наименование проверяемой организации", 400)
    f["org"] = find_org(org_block) or find_org(t)
    f["types"] = marked_audit_types(t)
    f["area"] = _segment(body, r"применительно к", r"область применения", 1500)
    f["basis"] = basis_refs(_segment(body, r"Основание", r"\n\s*3\s+[А-Я]", 400))
    f["sites"] = address_tokens(_segment(body, r"адрес аудитируемых", r"\n\s*5\s+[А-Я]", 800, next_line=True))
    f["okved"] = okved_codes(_segment(body, r"Подтвержденные коды", r"\n\s*6\s+[А-Я]", 400, next_line=True))
    block = exclusion_block(body)
    f["exclusions"] = exclusion_clauses(block) if block is not None else None
    f["compliance"] = _segment(body, r"Степень соответствия СМК критериям аудита", r"\n\s*15\s+[А-Я]", 200,
                               next_line=True).strip() or None
    f["period"] = find_period(_segment(t, r"Сроки проведения аудита", r"\n\s*\d{1,2}\s+[А-Я]", 200))
    f["team"] = parse_team(_segment(t, r"Состав экспертной группы", r"\n\s*9\s+[А-Я]", 600))
    rep = _segment(t, r"Уполномоченный представитель", r"Состав экспертной группы", 300)
    f["representative"] = find_persons(rep)[0] if find_persons(rep) else None

    # Подписи в конце: «Руководитель экспертной группы | | О.А. Сукорская | | 30.07.2026»
    f["date"], f["leader_sign"] = None, None
    for line in body.splitlines():
        if "руководитель экспертной группы" in line.lower():
            d = first_date(line)
            if d:
                f["date"] = d
                persons = find_persons(line)
                f["leader_sign"] = persons[0] if persons else None

    m = re.search(r"в ходе настоящего аудита выявлено\s+(\d+)", t, re.IGNORECASE)
    f["nc_23_2"] = int(m.group(1)) if m else None
    f["nc_total"] = f["nc_23_2"]

    def _nc(kind: str) -> Optional[int]:
        # «24 Результаты аудита»: значение стоит в ячейке ПЕРЕД названием строки
        mm = re.search(r"(?m)^\s*(\d*)\s*\|\s*Количество\s+" + kind + r"\s+несоответстви\w*(?!\s*,)", t,
                       re.IGNORECASE)
        if not mm:
            return None
        return int(mm.group(1)) if mm.group(1) else 0

    f["nc_significant"], f["nc_minor"] = _nc("значительных"), _nc("малозначительных")
    if f["nc_total"] is None and f["nc_significant"] is not None and f["nc_minor"] is not None:
        f["nc_total"] = f["nc_significant"] + f["nc_minor"]

    def _deadline(phrase: str) -> Optional[date]:
        seg = _segment(t, phrase, None, 600)
        m2 = re.search(r"\(?\s*до\s+(\d{1,2}\.\d{1,2}\.\d{2,4})", seg)
        return first_date(m2.group(1)) if m2 else None

    f["pkd_deadline"] = _deadline(r"представить\s+(?:его|план)|план корректирующих действий")
    f["kd_deadline"] = (_deadline(r"Сроки устранения несоответствий и вызвавших их причин")
                        or _deadline(r"отчет об устранении несоответствий"))
    f["obligation_pkd"] = bool(re.search(r"разработать план корректирующих действий", t, re.IGNORECASE))
    f["extend"] = is_marked(t, "расширить область")
    f["narrow"] = is_marked(t, "сузить область")

    # Список рассылки и подписи — в хвосте документа после «Список рассылки»
    i = body.lower().rfind("список рассылки")
    tail = body[i:] if i >= 0 else ""
    j = tail.lower().find("руководитель экспертной группы")
    f["distribution_org"] = find_org(tail[:j] if j >= 0 else tail)
    sign_block = tail[j:] if j >= 0 else ""
    k = sign_block.lower().find("представитель проверяемой организации")
    f["signers"] = find_persons(sign_block[:k] if k >= 0 else sign_block) if j >= 0 else []
    return f


# ---------- План аудита (П) ----------

def plan_fields(text: str) -> dict:
    """org (П1), area (П4 — область), basis (П2), sites (П3), period (П4 — сроки),
    representative (П5), team (П6), exclusions (П7), types (П9)."""
    t = text or ""
    body = t.split("=== ОТМЕТКИ")[0]
    i = t.lower().find("сокращенное наименование проверяемой организации")
    block = t[max(0, i - 400): i] if i >= 0 else t
    ex = exclusion_block(body)
    rep = _segment(body, r"Уполномоченный представитель", r"Состав экспертной группы", 300)
    return {
        "org": find_org(block, last=True) or find_org(t),
        "area": _segment(body, r"применительно к", r"область применения", 1500),
        "basis": basis_refs(_segment(body, r"Основание", r"\n\s*3\s+[А-Я]", 400)),
        "sites": address_tokens(_segment(body, r"адрес аудитируемых", r"\n\s*5\s+[А-Я]", 800, next_line=True)),
        "period": find_period(_segment(body, r"Сроки проведения аудита", r"\n\s*\d{1,2}\s+[А-Я]", 200)),
        "representative": find_persons(rep)[0] if find_persons(rep) else None,
        "team": parse_team(_segment(body, r"Состав экспертной группы", r"\n\s*8\s+[А-Я]", 600)),
        "exclusions": exclusion_clauses(ex) if ex is not None else None,
        "types": marked_audit_types(t),
    }


# ---------- Протоколы регистрации несоответствий (ПН) ----------

def pn_fields(text: str) -> dict:
    """protocols — по протоколу: {org, date, leader, has_sto}; version — версия формы."""
    t = text or ""
    starts = [m.start() for m in re.finditer(r"ПРОТОКОЛ\s+РЕГИСТРАЦИИ\s+НЕСООТВЕТСТВИ", t, re.IGNORECASE)]
    if not starts:
        starts = [0]
    chunks = [t[a: b] for a, b in zip(starts, starts[1:] + [len(t)])]
    protocols = []
    for ch in chunks:
        if "описание несоответствия" not in ch.lower():
            continue
        date_m = re.search(r"Дата\s*:?\s*(\d{1,2}\.\d{1,2}\.\d{2,4})", ch)
        leader_seg = _segment(ch, r"Руководитель экспертной группы", None, 250)
        persons = find_persons(leader_seg)
        protocols.append({
            "org": find_org(ch),
            "date": first_date(date_m.group(1)) if date_m else None,
            "leader": persons[0] if persons else None,
            "has_sto": bool(re.search(r"СТО\s+Газпром\s+9001", ch, re.IGNORECASE)),
        })
    vm = re.search(r"вер\.\s*\d+\s*от\s*[\d.]+", t)
    return {"protocols": protocols, "version": vm.group(0) if vm else None}


# ---------- ПКД и Отчет по КД (утверждаемые организацией документы) ----------

def count_table_rows(text: str) -> Optional[int]:
    """Строки таблицы несоответствий: после заголовка «№ п/п» — строки «N | …» с номерами
    1, 2, 3… подряд. Строка нумерации колонок «1 | 2 | 3 | …» под заголовком не считается.
    None — заголовка нет (например, плохой OCR)."""
    t = text or ""
    i = t.find("№ п/п")
    if i < 0:
        return None
    n = 0
    for m in re.finditer(r"(?m)^\s*(\d{1,3})\s*\|(.*)$", t[i:]):
        if re.fullmatch(r"(?:\s*\d{1,2}\s*\|?)+\s*", m.group(2)):
            continue  # «1 | 2 | 3 | 4 | 5» — номера колонок
        if int(m.group(1)) == n + 1:
            n += 1
    return n


def approved_doc_fields(text: str) -> dict:
    """Наименование организации (над «наименование проверяемой организации полностью»),
    дата утверждения (после «УТВЕРЖДАЮ»), число строк таблицы."""
    t = text or ""
    i = t.lower().find("наименование проверяемой организации полностью")
    block = t[max(0, i - 300): i] if i >= 0 else t[:1500]
    approve = _segment(t, r"УТВЕРЖДАЮ", None, 400)
    return {
        "org": find_org(block, last=True) or find_org(t),
        "date": first_date(approve),
        "rows": count_table_rows(t),
    }


# ---------- Листы оценки (ЛОО) ----------

_LOO_ROLES = [
    ("leader", "руководитель экспертной группы"),
    ("candidates", "кандидат"),
    ("experts", "член экспертной группы"),
    ("experts", "эксперт"),
]


def loo_fields(text: str) -> dict:
    """org (ЛОО1), period (ЛОО2), person (ЛОО3), roles (ЛОО31–33), date (ЛОО4),
    signer (ЛОО5), standards (ЛОО6), rating_rows (ЛОО7: подпись строки → число отметок)."""
    t = text or ""
    body = t.split("=== ОТМЕТКИ")[0]
    f = {"org": find_org(_segment(body, r"Организация", None, 300))}
    f["period"] = find_period(_segment(body, r"Дата аудита", None, 120))
    persons = find_persons(_segment(body, r"Оцениваемый", None, 200))
    f["person"] = persons[0] if persons else None

    signs = list(re.finditer(r"(?m)^\s*Дата\b(?!\s+аудита)", body))
    f["date"] = first_date(body[signs[-1].end(): signs[-1].end() + 80]) if signs else None
    signer = find_persons(_segment(body, r"Ф\.\s*И\.\s*О", None, 200))
    f["signer"] = signer[0] if signer else None

    marks = checkbox_marks(t)
    roles, standards, rows = [], [], {}
    for checked, label in marks:
        low = label.lower().replace("ё", "е")
        rows.setdefault(label, []).append(checked)
        if not checked:
            continue
        if re.search(r"гост|сто газпром|другое", low):
            standards.append(label)
    # Роль оцениваемого — отметки с одиночной подписью (у строк критериев по 5 клеток)
    for checked, label in marks:
        low = label.lower().replace("ё", "е")
        if checked and len(rows[label]) == 1:
            role = next((r for r, k in _LOO_ROLES if k in low), None)
            if role and role not in roles:
                roles.append(role)
    f["roles"] = roles
    f["standards"] = standards
    # Строки критериев — подписи, у которых не меньше трёх клеток оценки
    f["rating_rows"] = {label: sum(v) for label, v in rows.items() if len(v) >= 3}
    return f


# ---------- Разбивка ОКВЭД (ОБ), Приказ ЭГ (ПР), Договор (Д), Сертификат (Серт) ----------

def razbivka_fields(text: str) -> dict:
    """area (ОБ1 — «Заявленная область»), okved (ОБ2 — предлагаемые коды)."""
    t = text or ""
    codes_line = _segment(t, r"указать следующие коды", r"\n", 400)
    return {
        "area": _segment(t, r"Заявленная область", r"\n", 1500),
        "okved": okved_codes(codes_line) or okved_codes(t),
    }


def prikaz_fields(text: str) -> dict:
    """audit_type (ПР1), team (ПР2 — ФИО после «сформировать экспертную группу»)."""
    t = text or ""
    head = _segment(t, r"Для\s+проведения", r"ПРИКАЗЫВАЮ", 600) or t[:1500]
    return {
        "audit_type": main_audit_type(head),
        "team": find_persons(_segment(t, r"экспертную\s+группу", r"Руководитель\s+органа", 1200)),
    }


def contract_fields(text: str) -> dict:
    """org (Д2) — Заказчик из преамбулы: первая организация, кроме органа по сертификации;
    refs (Д1) — (номера, даты) из заголовка: «ДОГОВОР № … от …», у доп. соглашения —
    и свой номер, и «к Договору № … от …»."""
    t = text or ""
    head = t[:800]
    nums, dates = set(), set()
    for m in re.finditer(r"(?:ДОГОВОР|СОГЛАШЕНИЕ|Договору)\s*№\s*([A-Za-zА-Яа-я0-9/\-_.]+)", head, re.IGNORECASE):
        num = re.sub(r"[^0-9a-zа-я/]", "", m.group(1).lower()).strip("/")
        if num:
            nums.add(num)
        d = first_date(head[m.end(): m.end() + 120])
        if d:
            dates.add(d)
    return {"org": find_org(t[:4000]), "refs": (nums, dates)}


def sert_fields(text: str) -> dict:
    """valid_until (СЕРТ2 — предположительно окончание срока действия:
    «Срок действия с 28.07.2024 по 27.07.2027»)."""
    seg = _segment(text or "", r"Срок\s+действия", None, 120)
    m = re.search(r"по\s+(\d{1,2}\.\d{1,2}\.\d{2,4})", seg)
    return {"valid_until": first_date(m.group(1)) if m else None}


# ---------- Заключение об устранении несоответствий (З) ----------

def zakl_fields(text: str) -> dict:
    """org, approver (З1), approval_date, types (п.1), period, rows (п.3), leader (З2),
    sign_date, distribution_org (п.5) и отметки п.3–4."""
    t = text or ""
    body = t.split("=== ОТМЕТКИ")[0]
    approve = _segment(body, r"УТВЕРЖДАЮ", r"наименование проверяемой организации", 500)
    i = body.lower().find("наименование проверяемой организации")
    org_block = body[max(0, i - 300): i] if i >= 0 else body
    approvers = find_persons(approve)

    leader, sign_date = None, None
    for line in body.splitlines():
        if "руководитель экспертной группы" in line.lower() and first_date(line):
            persons = find_persons(line)
            leader, sign_date = (persons[0] if persons else None), first_date(line)

    dist = _segment(body, r"Список рассылки", r"Руководитель экспертной группы", 400)
    return {
        "org": find_org(org_block, last=True),
        "approver": approvers[0] if approvers else None,
        "approval_date": first_date(approve),
        "types": marked_audit_types(t),
        "period": find_period(_segment(body, r"Сроки проведения аудита", r"\n\s*\d\s+[А-Я]", 200)),
        "rows": count_table_rows(body),
        "leader": leader,
        "sign_date": sign_date,
        "distribution_org": find_org(dist),
        "eliminated": is_marked(t, "несоответствия устранены"),
        "implemented": is_marked(t, "корректирующие действия реализованы"),
        "issue": is_marked(t, "выдать сертификат", exact=True),
        "confirm": is_marked(t, "подтвердить действие сертификата", exact=True),
        "extend": is_marked(t, "расширить область"),
        "narrow": is_marked(t, "сузить область"),
    }


# ---------- 1 этап: Отчет по результатам 1 этапа (1О), Отчет об устранении (1Отчет),
# Заключение 1 этапа (1Закл) ----------

def otchet1_fields(text: str) -> dict:
    """org (1О1), approval_date (1О7), area (1О5), basis (1О6), period (1О8), team (1О9),
    nc_count (1О2 — «N | Указать количество несоответствий…»), nc_rows (строки таблицы
    1О10, п.10.1.8), conformity (п.10.1.4)."""
    t = text or ""
    body = t.split("=== ОТМЕТКИ")[0]
    approve = _segment(body, r"УТВЕРЖДАЮ", r"наименование проверяемой организации", 500)
    i = body.lower().find("наименование проверяемой организации")
    org_block = body[max(0, i - 300): i] if i >= 0 else body
    m = re.search(r"(\d+)\s*\|[^\n|]{0,20}Указать количество несоответствий", body, re.IGNORECASE)
    nc_table = _segment(body, r"Описание несоответствий по результатам анализа документации",
                        r"\n\s*10\.2\s+[А-Я]", 20000)
    return {
        "org": find_org(org_block, last=True),
        "approval_date": first_date(approve),
        "area": _segment(body, r"применительно к", r"область применения", 1500),
        "basis": basis_refs(_segment(body, r"Основание", r"\n\s*3\s+[А-Я]", 400)),
        "period": find_period(_segment(body, r"Сроки проведения аудита", r"\n\s*\d{1,2}\s+[А-Я]", 200)),
        "team": parse_team(_segment(body, r"Состав экспертной группы", r"\n\s*8\s+[А-Я]", 600)),
        "nc_count": int(m.group(1)) if m else None,
        "nc_rows": count_table_rows(nc_table) if nc_table else None,
        "conformity": _segment(body, r"10\.1\.4\s+Соответствие документации", r"\n\s*10\.1\.5", 400,
                               next_line=True).strip() or None,
    }


def zakl1_fields(text: str) -> dict:
    """Заключение 1 этапа — как Заключение (zakl_fields) плюс отметка «Проведение второго
    этапа … аудита возможно». Образца нет: разбор — по аналогии с Заключением того же ОС."""
    f = zakl_fields(text)
    f["second_stage"] = is_marked(text, "проведение второго этапа")
    return f
