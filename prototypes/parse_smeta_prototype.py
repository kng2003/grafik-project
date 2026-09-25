# -*- coding: utf-8 -*-
"""
Прототип парсера локальной сметы (ресурсный метод, ГРАНД-Смета) в структуру
для построения графика работ (Ганта).

Реализует три проверенные на реальных примерах гипотезы:
 1. Консолидация позиций-довесков ("к норме/нормам ...") к ближайшей
    предшествующей позиции с совпадающим базовым шифром в той же подгруппе.
 2. Отделение позиций-расценок (ГЭСН/ФЕР/ФССЦ/ТЕР — реальная работа) от
    позиций-цен (числовой код без префикса — просто цена ресурса из ССЦ),
    последние не являются отдельными задачами графика, а прикрепляются к
    ближайшей предыдущей позиции-работе как её материал/ресурс.
 3. Определение ведущего ресурса: сравнение суммарных чел-ч (группа ОТ(ЗТ))
    и суммарных маш-ч (группа ЭМ). Если ручной труд больше или равен машинному
    — ведущий ресурс трудовой (работа считается ручной/условно-механизированной).
    Если машинное время больше — ведущий ресурс станок/машина с максимальным
    временем внутри группы ЭМ. Если внутри группы ЭМ больше одной машины —
    результат помечается на ручную проверку (это тот самый случай, где
    правило "макс. часов" не совпадает с наименованием расценки по техчасти).

Ограничения (сознательно, MVP): без доступа к технической части ГЭСН/ФЕР
нельзя получить название профессии ведущего рабочего и точный состав звена —
парсер даёт только структурный ярлык ("Рабочие, средний разряд X" или имя
машины из каталога ресурсов).
"""
import openpyxl
import re
import sys
import json

CODE_RE = re.compile(r'^(ФЕР|ГЭСН|ФССЦ|ТЕР)[а-я]{0,2}[\d.\-]+', re.IGNORECASE)
ADDON_RE = re.compile(r'к\s+норм(?:е|ам|ы)?\s+([0-9.\-,\s]+)', re.IGNORECASE)
SECTION_RE = re.compile(r'^Раздел\s+\d+\.', re.IGNORECASE)


def cell(ws, r, c):
    return ws.cell(row=r, column=c).value


def norm_code(raw):
    if not raw:
        return None
    return str(raw).split('\n')[0].strip()


def base_prefix(code):
    """Убирает последний '-NN' сегмент шифра, чтобы сравнивать базовые нормы."""
    parts = code.split('-')
    if len(parts) <= 1:
        return code
    return '-'.join(parts[:-1])


def parse_smeta(path, sheet_name=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    # 1. Собираем номера позиций и границы блоков + секции/подгруппы.
    pos_rows = []  # (row, pos_label)
    section = None
    subgroup1 = None
    subgroup2 = None
    context_by_row = {}
    started = False  # включаем сбор позиций только после первого "Раздел N."
    for r in range(1, ws.max_row + 1):
        c1 = cell(ws, r, 1)
        if c1 is None:
            continue
        s = str(c1).strip()
        if SECTION_RE.match(s):
            section = s
            subgroup1 = None
            subgroup2 = None
            started = True
            continue
        if not started:
            continue
        # текстовые подзаголовки подгрупп: короткая строка без цифр в начале
        if not re.match(r'^\d', s) and cell(ws, r, 2) is None and cell(ws, r, 3) is None:
            if subgroup1 is None:
                subgroup1 = s
            else:
                subgroup2 = s
            continue
        m = re.match(r'^(\d+)', s)
        if m:
            pos_rows.append((r, m.group(1)))
            context_by_row[r] = (section, subgroup1, subgroup2)

    positions = []
    prev_work_pos = None  # для прикрепления PRICE_REF и поиска базовой нормы

    for idx, (row_start, pos_label) in enumerate(pos_rows):
        row_end = pos_rows[idx + 1][0] if idx + 1 < len(pos_rows) else ws.max_row + 1
        code_raw = cell(ws, row_start, 2)
        code = norm_code(code_raw)
        name = cell(ws, row_start, 3)
        name = str(name) if name is not None else ''
        unit = cell(ws, row_start, 8)
        qty = cell(ws, row_start, 11) if cell(ws, row_start, 11) is not None else cell(ws, row_start, 9)
        section, sg1, sg2 = context_by_row.get(row_start, (None, None, None))

        is_work = bool(code) and bool(CODE_RE.match(code))

        pos = {
            'pos': pos_label,
            'row': row_start,
            'code': code,
            'name': name.replace('\n', ' ').strip(),
            'unit': unit,
            'qty': qty,
            'section': section,
            'subgroup1': sg1,
            'subgroup2': sg2,
            'type': None,
            'addon_of': None,
            'attached_price_refs': [],
            'labor_hours': 0.0,
            'machine_hours': 0.0,
            'machines': [],
            'materials_embedded': [],
            'leading_resource': None,
            'leading_confidence': None,
            'review': [],
        }

        if not is_work:
            pos['type'] = 'PRICE_REF'
            if prev_work_pos is not None:
                prev_work_pos['attached_price_refs'].append({
                    'pos': pos_label, 'name': pos['name'], 'unit': unit, 'qty': qty,
                })
            else:
                pos['review'].append('нет предшествующей работы для прикрепления цены')
            positions.append(pos)
            continue

        pos['type'] = 'WORK'

        # проверка на позицию-довесок ("к норме ...")
        addon_match = ADDON_RE.search(name)
        if addon_match:
            ref_raw = addon_match.group(1)
            ref_codes = re.findall(r'[\d]+(?:-[\d]+)*', ref_raw)
            ref_prefixes = set()
            for rc in ref_codes:
                ref_prefixes.add(rc)
            # ищем ближайшую предыдущую WORK-позицию, чей код содержит один из ref_prefixes
            found = None
            for p in reversed(positions):
                if p['type'] != 'WORK':
                    continue
                if p['subgroup1'] != sg1 or p['subgroup2'] != sg2:
                    # выходим за пределы текущей технологической подгруппы
                    break
                code_num = re.sub(r'^[^\d]+', '', p['code']).replace('.', '-')
                for rp in ref_prefixes:
                    if rp.replace('.', '-').strip() == code_num.strip():
                        found = p
                        break
                if found:
                    break
            if found:
                pos['addon_of'] = found['pos']
            else:
                pos['review'].append(f'довесок "{ref_raw.strip()}" — базовая позиция не найдена рядом')

        # разбор ресурсных групп внутри блока
        r = row_start + 1
        current_group = None
        while r < row_end:
            v2 = cell(ws, r, 2)
            v3 = cell(ws, r, 3)
            v8 = cell(ws, r, 8)
            v9 = cell(ws, r, 9)
            v11 = cell(ws, r, 11)

            # маркер группы: короткий числовой индекс в col2 + метка в col3
            if isinstance(v2, str) and v2.strip().isdigit() and len(v2.strip()) <= 2 and isinstance(v3, str):
                label = v3.strip()
                if label in ('ОТ(ЗТ)', 'ЭМ', 'М', 'ЗТ'):
                    current_group = label
                    if label == 'ОТ(ЗТ)' and v11 is not None:
                        pos['labor_hours'] += float(v11)
                    r += 1
                    continue

            # ресурсная строка (каталожный шифр ресурса + единица измерения)
            if isinstance(v2, str) and v8 and current_group:
                qv = v11 if v11 is not None else v9
                try:
                    qv = float(qv) if qv is not None else 0.0
                except (TypeError, ValueError):
                    qv = 0.0
                unit_s = str(v8).strip()
                if current_group == 'ЭМ' and unit_s == 'маш.-ч':
                    pos['machines'].append({'code': v2.strip(), 'name': (v3 or '').replace('\n', ' '), 'hours': qv})
                    pos['machine_hours'] += qv
                elif current_group == 'ЭМ' and unit_s == 'чел.-ч':
                    # это труд машиниста, привязанный к машине — не отдельный ведущий ресурс
                    pass
                elif current_group == 'ОТ(ЗТ)' and unit_s == 'чел.-ч':
                    pass  # уже учли по итоговой строке группы
                elif current_group == 'М':
                    pos['materials_embedded'].append({'code': v2.strip(), 'name': (v3 or '').replace('\n', ' '), 'unit': unit_s, 'qty': qv})
            r += 1

        # определение ведущего ресурса
        L = pos['labor_hours']
        M = pos['machine_hours']
        if M == 0 and L == 0:
            pos['leading_resource'] = 'не определено (нет ресурсных строк)'
            pos['leading_confidence'] = 'Догадка'
            pos['review'].append('нет данных по ресурсам')
        elif M == 0:
            pos['leading_resource'] = f'Рабочие (труд), {L:.1f} чел-ч'
            pos['leading_confidence'] = 'Скорее всего'
        elif L == 0:
            top = max(pos['machines'], key=lambda x: x['hours']) if pos['machines'] else None
            pos['leading_resource'] = f"{top['name']} ({top['hours']:.1f} маш-ч)" if top else 'машина не определена'
            pos['leading_confidence'] = 'Скорее всего' if len(pos['machines']) <= 1 else 'Догадка'
            if pos['leading_confidence'] == 'Догадка':
                pos['review'].append('несколько машин без ручного труда, максимум маш-ч не даёт явного лидера — проверить по техчасти')
        else:
            if L >= M:
                pos['leading_resource'] = f'Рабочие (труд), {L:.1f} чел-ч (механизация вспомогательная, {M:.1f} маш-ч)'
                pos['leading_confidence'] = 'Скорее всего'
                # труд уже уверенно ведущий ресурс — несколько вспомогательных машин это не меняют,
                # ручная проверка не нужна (было: флаг ставился всегда при >1 машине, что переоценивало
                # долю позиций на проверку — см. находку от 2026-09-03 на объекте Сокур)
            else:
                top = max(pos['machines'], key=lambda x: x['hours']) if pos['machines'] else None
                pos['leading_resource'] = f"{top['name']} ({top['hours']:.1f} маш-ч; труд {L:.1f} чел-ч)" if top else 'машина не определена'
                pos['leading_confidence'] = 'Скорее всего' if len(pos['machines']) <= 1 else 'Догадка'
                if pos['leading_confidence'] == 'Догадка':
                    pos['review'].append('машина ведущая, но несколько машин без явного максимума — сверить с наименованием по техчасти ГЭСН/ФЕР')

        positions.append(pos)
        prev_work_pos = pos

    return positions


def write_report(positions, out_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = 'Позиции сметы'
    headers = [
        '№ п/п', 'Раздел', 'Подгруппа 1', 'Подгруппа 2', 'Тип', 'Шифр',
        'Наименование', 'Ед. изм.', 'Кол-во', 'Ведущий ресурс',
        'Уверенность', 'Прикреплено (материалы/цены)', 'На проверку',
    ]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)

    review_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    addon_fill = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid')

    for p in positions:
        if p['type'] == 'PRICE_REF':
            type_label = 'Цена-ресурс (не задача графика)'
        elif p['addon_of']:
            type_label = f'Довесок к поз. {p["addon_of"]}'
        else:
            type_label = 'Работа (задача графика)'

        attached = '; '.join(
            f"{a['name']} ({a['qty']} {a['unit']})" for a in p['attached_price_refs']
        )
        row = [
            p['pos'], p['section'], p['subgroup1'], p['subgroup2'], type_label,
            p['code'], p['name'], p['unit'], p['qty'],
            p['leading_resource'], p['leading_confidence'], attached,
            '; '.join(p['review']),
        ]
        ws.append(row)
        r = ws.max_row
        if p['review']:
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = review_fill
        elif p['addon_of']:
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = addon_fill

    widths = [8, 20, 18, 18, 26, 20, 45, 10, 10, 32, 14, 40, 35]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

    # сводный лист
    ws2 = wb.create_sheet('Сводка')
    total = len(positions)
    work = [p for p in positions if p['type'] == 'WORK']
    addons = [p for p in work if p['addon_of']]
    schedule_tasks = [p for p in work if not p['addon_of']]
    price_refs = [p for p in positions if p['type'] == 'PRICE_REF']
    need_review = [p for p in positions if p['review']]
    ws2.append(['Показатель', 'Значение'])
    ws2.append(['Всего позиций в смете', total])
    ws2.append(['Из них — расценки (работы), ГЭСН/ФЕР/ФССЦ/ТЕР', len(work)])
    ws2.append(['  в т.ч. довески (консолидированы к базовой позиции)', len(addons)])
    ws2.append(['  в т.ч. самостоятельные задачи графика', len(schedule_tasks)])
    ws2.append(['Позиции-цены (не задачи графика, прикреплены к работе)', len(price_refs)])
    ws2.append(['Позиций, требующих ручной проверки', len(need_review)])
    for c in (1, 2):
        ws2.cell(row=1, column=c).font = Font(bold=True)
    ws2.column_dimensions['A'].width = 55
    ws2.column_dimensions['B'].width = 12

    wb.save(out_path)


if __name__ == '__main__':
    path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else None
    result = parse_smeta(path)
    if out_path:
        write_report(result, out_path)
        print(f'Отчёт сохранён: {out_path}')
    else:
        print(json.dumps(result, ensure_ascii=False, indent=1))
