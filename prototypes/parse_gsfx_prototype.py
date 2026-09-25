# -*- coding: utf-8 -*-
"""
Прототип парсера нативного формата Гранд-Сметы (.gsfx) в структуру для
построения графика работ. .gsfx — это ZIP-архив (Properties.txt, Data.xml,
Data.sign), Data.xml — обычный XML в кодировке windows-1251 (без шифрования,
Crypted=False у всех проверенных файлов).

В отличие от Excel-экспорта (см. parse_smeta_prototype.py), этот формат даёт
структуру напрямую, без угадывания по тексту/положению ячеек:
 - каждая расценка — тег <Position Code Caption Quantity Units>;
 - её ресурсы типизированы: <Tzr> труд рабочих (чел-ч), <Tzm> труд
   машинистов (чел-ч, обслуживание техники — не самостоятельный ресурс),
   <Mch> машины (маш-ч), <Mat> материалы;
 - материалы уже вложены в свою позицию-работу — гипотеза "позиции-цены
   прикрепляются к ближайшей предыдущей работе" здесь не нужна;
 - заголовки подгрупп — отдельный тег <Header>, отличимый от обычных
   примечаний <Comment> — не нужно гадать по стилю ячейки.

Оставшиеся гипотезы (проверены здесь так же, как в Excel-версии):
 1. Консолидация довесков — по тексту Caption ("к норме/нормам ...") в
    пределах одного Chapter, к ближайшей предыдущей позиции.
 2. Ведущий ресурс — сумма Tzr против суммы Mch; при равенстве нулю трудовых
    ресурсов и >1 машины — тай-брейк максимум маш-ч, флаг "на проверку"
    только когда confidence реально "Догадка" (см. находку 2026-09-03 в
    parse_smeta_prototype.py — тот же баг был исправлен и здесь сразу).
"""
import zipfile
import re
import sys
import json
import xml.etree.ElementTree as ET

ADDON_RE = re.compile(r'к\s+норм(?:е|ам|ы)?\s+([0-9.\-,\s]+)', re.IGNORECASE)
CODE_RE = re.compile(r'^(ФЕР|ГЭСН|ФССЦ|ТЕР)[а-я]{0,2}[\d.\-]+', re.IGNORECASE)


def to_float(s):
    if s is None:
        return 0.0
    s = str(s).replace('\xa0', '').replace(' ', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def base_prefix(code):
    parts = code.split('-')
    if len(parts) <= 1:
        return code
    return '-'.join(parts[:-1])


def load_data_xml(path):
    with zipfile.ZipFile(path) as z:
        raw = z.read('Data.xml')
    # заголовок сообщает кодировку явно (обычно windows-1251)
    m = re.search(rb'encoding="([\w-]+)"', raw[:200])
    enc = m.group(1).decode('ascii') if m else 'windows-1251'
    text = raw.decode(enc)
    return ET.fromstring(text)


def parse_gsfx(path):
    root = load_data_xml(path)
    positions = []
    prev_by_chapter = {}  # chapter_sysid -> список последних позиций (для консолидации)

    for chapter in root.iter('Chapter'):
        chapter_caption = chapter.get('Caption')
        chapter_id = chapter.get('SysID')
        headers = []  # стек текущих заголовков подгрупп в этом разделе
        prev_pos = None

        for child in chapter:
            if child.tag == 'Header':
                headers.append(child.get('Caption'))
                continue
            if child.tag != 'Position':
                continue

            code = child.get('Code') or ''
            caption = child.get('Caption') or ''
            units = child.get('Units')
            qty_raw = child.get('Quantity')

            res = child.find('Resources')
            tzr = 0.0
            mch_total = 0.0
            machines = []
            materials = []
            if res is not None:
                for t in res.findall('Tzr'):
                    tzr += to_float(t.get('Quantity'))
                for m in res.findall('Mch'):
                    q = to_float(m.get('Quantity'))
                    mch_total += q
                    machines.append({'code': m.get('Code'), 'name': m.get('Caption'), 'hours': q})
                for mat in res.findall('Mat'):
                    materials.append({
                        'code': mat.get('Code'), 'name': mat.get('Caption'),
                        'unit': mat.get('Units'), 'qty': to_float(mat.get('Quantity')),
                    })

            # НАХОДКА 2026-09-03 (исправление собственной ошибки): не все материалы/цены
            # вложены в Resources родительской работы — часть встречается как ОТДЕЛЬНЫЕ
            # top-level <Position> без кода ГЭСН/ФЕР/ФССЦ/ТЕР (числовой код ценника или
            # произвольный текст вроде "Цены приняты по замечаниям..."), совсем без
            # ресурсов (Tzr/Mch пустые) — это тот же случай "позиция-цена", что и в
            # Excel-версии, просто в .gsfx он тоже не полностью снят структурой документа.
            is_price_ref = not CODE_RE.match(code)

            pos = {
                'chapter': chapter_caption,
                'subgroup': ' / '.join(headers[-2:]) if headers else None,
                'code': code,
                'name': caption,
                'unit': units,
                'qty_raw': qty_raw,
                'type': 'PRICE_REF' if is_price_ref else 'WORK',
                'labor_hours': tzr,
                'machine_hours': mch_total,
                'machines': machines,
                'materials': materials,
                'addon_of': None,
                'attached_to': None,
                'leading_resource': None,
                'leading_confidence': None,
                'review': [],
            }

            if is_price_ref:
                if prev_pos is not None and prev_pos['type'] == 'WORK':
                    pos['attached_to'] = prev_pos['code']
                else:
                    pos['review'].append('позиция-цена без предшествующей работы для прикрепления')
                positions.append(pos)
                # позиция-цена не становится "prev_pos" для консолидации довесков —
                # довески консолидируются только относительно предыдущей РАБОТЫ
                continue

            m_addon = ADDON_RE.search(caption)
            if m_addon and prev_pos is not None and prev_pos['type'] == 'WORK':
                ref_codes = re.split(r'[,;]|\s+и\s+', m_addon.group(1))
                ref_prefixes = {c.strip().rstrip('.') for c in ref_codes if c.strip()}
                prev_prefix = base_prefix(prev_pos['code'])
                cur_prefix = base_prefix(code)
                if cur_prefix == base_prefix(prev_pos['code']) or any(
                    prev_prefix.endswith(rp) or rp in prev_prefix for rp in ref_prefixes
                ):
                    pos['addon_of'] = prev_pos['code']

            L, M = tzr, mch_total
            if M == 0 and L == 0:
                pos['leading_resource'] = 'не определено (нет ресурсных строк)'
                pos['leading_confidence'] = 'Догадка'
                pos['review'].append('нет данных по ресурсам')
            elif L == 0:
                top = max(machines, key=lambda x: x['hours']) if machines else None
                pos['leading_resource'] = f"{top['name']} ({top['hours']:.2f} маш-ч)" if top else 'машина не определена'
                pos['leading_confidence'] = 'Скорее всего' if len(machines) <= 1 else 'Догадка'
                if pos['leading_confidence'] == 'Догадка':
                    pos['review'].append('несколько машин без ручного труда, максимум маш-ч не даёт явного лидера')
            else:
                if L >= M:
                    pos['leading_resource'] = f'Рабочие (труд), {L:.2f} чел-ч (механизация вспомогательная, {M:.2f} маш-ч)'
                    pos['leading_confidence'] = 'Скорее всего'
                else:
                    top = max(machines, key=lambda x: x['hours']) if machines else None
                    pos['leading_resource'] = f"{top['name']} ({top['hours']:.2f} маш-ч; труд {L:.2f} чел-ч)" if top else 'машина не определена'
                    pos['leading_confidence'] = 'Скорее всего' if len(machines) <= 1 else 'Догадка'
                    if pos['leading_confidence'] == 'Догадка':
                        pos['review'].append('машина ведущая, но несколько машин без явного максимума')

            positions.append(pos)
            prev_pos = pos

    return positions


def write_report(positions, out_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = 'Позиции сметы'
    headers = [
        '№', 'Раздел', 'Подгруппа', 'Тип', 'Шифр', 'Наименование', 'Ед. изм.',
        'Кол-во (формула)', 'Ведущий ресурс / прикреплено к', 'Уверенность', 'Материалы (вложены)', 'На проверку',
    ]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)

    review_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    addon_fill = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid')
    price_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')

    for i, p in enumerate(positions, start=1):
        mats = '; '.join(f"{m['name']} ({m['qty']:.3g} {m['unit']})" for m in p['materials'][:5])
        if len(p['materials']) > 5:
            mats += f" … +{len(p['materials'])-5}"
        if p['type'] == 'PRICE_REF':
            type_label = 'Цена-ресурс (не задача графика)'
            lead = f"прикреплено к {p['attached_to']}" if p['attached_to'] else '—'
        elif p['addon_of']:
            type_label = f'Довесок к поз. {p["addon_of"]}'
            lead = p['leading_resource']
        else:
            type_label = 'Работа (задача графика)'
            lead = p['leading_resource']
        row = [
            i, p['chapter'], p['subgroup'], type_label, p['code'], p['name'], p['unit'],
            p['qty_raw'], lead, p['leading_confidence'], mats,
            '; '.join(p['review']),
        ]
        ws.append(row)
        r = ws.max_row
        if p['review']:
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = review_fill
        elif p['type'] == 'PRICE_REF':
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = price_fill
        elif p['addon_of']:
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = addon_fill

    widths = [5, 24, 24, 26, 20, 45, 8, 22, 30, 14, 45, 35]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

    ws2 = wb.create_sheet('Сводка')
    total = len(positions)
    work = [p for p in positions if p['type'] == 'WORK']
    price_refs = [p for p in positions if p['type'] == 'PRICE_REF']
    addons = [p for p in work if p['addon_of']]
    tasks = [p for p in work if not p['addon_of']]
    need_review = [p for p in positions if p['review']]
    ws2.append(['Показатель', 'Значение'])
    ws2.append(['Всего позиций в смете', total])
    ws2.append(['Из них — расценки (работы)', len(work)])
    ws2.append(['  в т.ч. довески (консолидированы)', len(addons)])
    ws2.append(['  в т.ч. самостоятельные задачи графика', len(tasks)])
    ws2.append(['Позиции-цены (прикреплены к работе)', len(price_refs)])
    ws2.append(['Позиций, требующих ручной проверки', len(need_review)])
    for c in (1, 2):
        ws2.cell(row=1, column=c).font = Font(bold=True)
    ws2.column_dimensions['A'].width = 45
    ws2.column_dimensions['B'].width = 12

    wb.save(out_path)


if __name__ == '__main__':
    path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else None
    result = parse_gsfx(path)
    if out_path:
        write_report(result, out_path)
        print(f'Отчёт сохранён: {out_path}')
    else:
        print(json.dumps(result, ensure_ascii=False, indent=1))
