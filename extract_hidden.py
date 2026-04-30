#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Extract customXML, styles, and all hidden data from the docx
"""

import zipfile
import json

DOC_PATH = r"D:\Perpektiva\Пакет 2\ИИ -ЧК -План АУДИТА.docx"

with zipfile.ZipFile(DOC_PATH, 'r') as z:
    # Custom XML
    print("=" * 80)
    print("CUSTOM XML (пользовательские данные)")
    print("=" * 80)
    
    if 'customXML/item1.xml' in z.namelist():
        item1 = z.read('customXML/item1.xml').decode('utf-8')
        print(item1)
    
    if 'customXML/itemProps1.xml' in z.namelist():
        props = z.read('customXML/itemProps1.xml').decode('utf-8')
        print("\nItem Props:")
        print(props)
    
    # Settings
    print("\n" + "=" * 80)
    print("SETTINGS")
    print("=" * 80)
    
    if 'word/settings.xml' in z.namelist():
        settings = z.read('word/settings.xml').decode('utf-8')
        print(settings[:2000])
    
    # Styles (first 3000 chars)
    print("\n" + "=" * 80)
    print("STYLES (фрагмент)")
    print("=" * 80)
    
    if 'word/styles.xml' in z.namelist():
        styles = z.read('word/styles.xml').decode('utf-8')
        print(styles[:3000])
    
    # Document XML - search for field markers, bookmarks, hyperlinks
    print("\n" + "=" * 80)
    print("ПОИСК BOOKMARK, HYPERLINK, и других элементов")
    print("=" * 80)
    
    import xml.etree.ElementTree as ET
    doc_xml = z.read('word/document.xml').decode('utf-8')
    
    # Search for interesting patterns
    patterns = [
        'bookmarkStart', 'bookmarkEnd',
        'hyperlink',
        'comment',
        'instrText',  # field codes
        'FORMTEXT',   # form text field
        'FORMCHECKBOX', # form checkbox
        'FORMDROPDOWN', # form dropdown
        'checked',
        'w:tick',
        'R', '☐', '☑', '✓', '√', '□', '■',  # checkbox characters
        'ИИ1', 'ИИ2', 'ИИ3', 'ИИ4', 'ИИ5', 'ИИ6', 'ИИ7', 'ИИ8', 'ИИ9', 'ИИ10', 'ИИ11', 'ИИ12',
    ]
    
    for pattern in patterns:
        count = doc_xml.count(pattern)
        if count > 0:
            # Find context
            positions = []
            start = 0
            while True:
                idx = doc_xml.find(pattern, start)
                if idx == -1:
                    break
                context_start = max(0, idx - 50)
                context_end = min(len(doc_xml), idx + 100)
                context = doc_xml[context_start:context_end].replace('\n', ' ')
                positions.append(f"pos={idx}: ...{context}...")
                start = idx + 1
                if len(positions) >= 3:
                    break
            
            print(f"\n  '{pattern}' найдено: {count} раз")
            for p in positions[:2]:
                print(f"    {p}")
