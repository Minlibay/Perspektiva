import pytest

import doc_types

# Начала текстов — по образцам «ЧК ИИ/18.09» (без данных клиентов)
HEADS = [
    ("=== КОЛОНТИТУЛ (header) ===\nОС СМК | План аудита СМК\nПЛАН\nаудита системы менеджмента качества", "plan"),
    ("Орган по сертификации СМК\nП Р И К А З\n№ П/13-07/01 13 июля 2026 г.\nО назначении экспертной группы", "prikaz"),
    ("Перечень кодов ОКВЭД 2, соответствующих видам продукции", "razbivka"),
    ("ДОГОВОР № 28/01-01-2026\nг. Москва", "dogovor"),
    ("ДОПОЛНИТЕЛЬНОЕ СОГЛАШЕНИЕ № 3 к Договору № ИК05/06-2024", "dop_soglashenie"),
    ("=== Лист: Лист1 ===\nНазвание Заказчика\n№ | Показатель | Информация на дату предыдущего аудита", "info"),
    ("ПРИЛОЖЕНИЕ к сертификату соответствия № ОГН1.RU.1416", "sertifikat"),
    ("Орган по сертификации СМК\nРАСЧЕТ ТРУДОЕМКОСТИ (ПРОДОЛЖИТЕЛЬНОСТИ) СЕРТИФИКАЦИИ", "trudoemkost"),
    ("З А Я В К А\nна проведение добровольной сертификации", "zayavka"),
    ("ОС СМК | Отчет по результатам первого этапа ресертификационного аудита СМК", "otchet1"),
    ("УТВЕРЖДАЮ\nОТЧЕТ об устранении несоответствий, выявленных при анализе документации "
     "в процессе первого этапа", "otchet_ustr1"),
    ("УТВЕРЖДАЮ\nОТЧЕТ об устранении несоответствий, выявленных при аудите системы", "otchet_kd"),
    ("ОС СМК | Акт по результатам аудита СМК\nАКТ № 1 по результатам аудита", "akt"),
    ("ОС СМК | ПРОТОКОЛ РЕГИСТРАЦИИ НЕСООТВЕТСТВИЯ | Акт № 1", "pn"),
    ("Лист оценки эксперта проверяемой организацией", "loo_org"),
    ("Лист оценки экспертов друг другом внутри экспертной группы", "loo_exp"),
    ("ПЛАН корректирующих действий по несоответствиям", "pkd"),
    ("ЗАКЛЮЧЕНИЕ о результатах анализа сведений о выполнении корректирующих действий", "zakl"),
    ("СВОДНЫЙ АКТ ИССЛЕДОВАНИЯ СИСТЕМЫ МЕНЕДЖМЕНТА КАЧЕСТВА", "svod"),
]


@pytest.mark.parametrize("head,expected", HEADS)
def test_classify_by_content(head, expected):
    assert doc_types.classify_by_content(head) == expected


def test_type_named_only_far_in_text_is_ignored():
    # Ссылка на договор в глубине Плана не делает План договором
    text = "ПЛАН аудита системы менеджмента" + " x" * 1000 + " по Договору № 1"
    assert doc_types.classify_by_content(text) == "plan"


@pytest.mark.parametrize("name,expected", [
    ("3.1 Договор №05-09 от 03.09.25.pdf", "dogovor"),
    ("контракт 12.pdf", "dogovor"),          # «акт» внутри слова не считается Актом
    ("ДС2 СМК ВИК(расширение СМК).pdf", "dop_soglashenie"),
    ("4. Приказ об ЭГ_Ромашка.pdf", "prikaz"),
    ("ИИ Отчет по КД 2 этап.pdf", "otchet_kd"),
    ("Бланк заявки для СМК.pdf", "zayavka"),
    ("Скан.pdf", None),
])
def test_classify_by_name(name, expected):
    assert doc_types.classify_by_name(name) == expected


def test_content_wins_over_name():
    assert doc_types.classify_document("Договор.pdf", "ПРИКАЗ\nО назначении экспертной группы") == "prikaz"
    assert doc_types.classify_document("Скан.pdf", "") == "unknown"


def test_samples_classified(samples_dir):
    import main

    expected = {
        "ИИ Акт инспекционный контроль.docx": "akt",
        "ИИ Закл.docx": "zakl",
        "ИИ Лист Оценки  организацией.docx": "loo_org",
        "ИИ Лист оценки экспертами.docx": "loo_exp",
        "ИИ Отчет 1 этап РА УТЕКО .docx": "otchet1",
        "ИИ Отчет об устранении 1 этап НС.pdf": "otchet_ustr1",
        "ИИ План аудита.docm": "plan",
        "ИИ План КД.docx": "pkd",
        "ИИ Приказ об ЭГ.pdf": "prikaz",
        "ИИ Протоколы несоотв.docx": "pn",
        "ИИ СВОДНЫЙ АКТ.docx": "svod",
        "ИИ- Трудоемкость.docx": "trudoemkost",
        "Информация для подготовки к аудиту (2).xlsx": "info",
    }
    for name, dtype in expected.items():
        path = samples_dir / name
        suffix = path.suffix.lower()
        if suffix in (".docx", ".docm"):
            text = main.extract_text_from_docx(str(path))
        elif suffix == ".pdf":
            text = main.extract_text_from_pdf(str(path))
        else:
            text = main.extract_text_from_xlsx(str(path))
        assert doc_types.classify_document(name, text) == dtype, name
