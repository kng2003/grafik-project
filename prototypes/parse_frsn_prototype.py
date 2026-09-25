# -*- coding: utf-8 -*-
import re, sys, glob

SHIFR_START = re.compile(r'^Н-\d+-\d+-\d+')
FULL_SHIFR = re.compile(r'^Н-\d+-\d+-\d+-\d+-\d+$')
BARE_SUFFIX = re.compile(r'^\d+-\d+$')

def clean(s):
    s = s or ''
    s = s.replace('<br>', ' ')
    s = re.sub(r'\*\*', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def split_row(line):
    line = line.rstrip('\n')
    if not line.startswith('|'):
        return None
    parts = line.split('|')
    if parts and parts[0] == '':
        parts = parts[1:]
    if parts and parts[-1] == '':
        parts = parts[:-1]
    return [p.strip() for p in parts]

def is_separator_row(cells):
    return all(re.fullmatch(r':?-{2,}:?', c) for c in cells if c != '')

def is_numeric_header_row(cells):
    return bool(cells) and bool(re.fullmatch(r'\*\*1\*\*', cells[0]))

def is_header_row(cells):
    joined = ' '.join(cells)
    return ('Шифр' in joined and 'Наименование' in joined) or is_numeric_header_row(cells)

def is_noise_row(cells):
    joined = ''.join(cells)
    if 'Табли' in joined:
        return True
    if 'остав звена рабочих' in joined or 'Состав звена рабочих' in joined:
        return True
    non_empty = [c for c in cells if c.strip() not in ('', '-')]
    if non_empty and all(c.strip().startswith('**') for c in non_empty) and len(non_empty) <= 4:
        return True
    return False

def cell(cells, i):
    return cells[i] if len(cells) > i else ''

def detect_layout(cells):
    # cells for a numeric header row, e.g. ['**1**','**2**','**3**','**4**','**5**','**6**','**7**','**8**','**9**']
    # or merged variant: [...,'**6**<br>**7**', '**8**','**9**']  (8 cells total)
    if len(cells) >= 9:
        return 'std9'
    if len(cells) == 8:
        return 'merged67'
    return 'std9'  # fallback, unknown -> assume standard

def new_record(lineno):
    return {
        'shifr_frags': [], 'name_parts': [], 'unit_parts': [], 'work_parts': [],
        'lab_parts': [], 'mach_parts': [], 'crew_rows': [], 'first_line': lineno,
    }

def add_start(rec, cells, layout):
    rec['shifr_frags'] += [x for x in cells[0].split('<br>') if x]
    rec['name_parts'].append(cell(cells, 1))
    rec['unit_parts'].append(cell(cells, 2))
    rec['work_parts'].append(cell(cells, 3))
    rec['lab_parts'].append(cell(cells, 4))
    if layout == 'std9':
        rec['mach_parts'].append(cell(cells, 5))
        rec['crew_rows'].append((cell(cells, 6), cell(cells, 7), cell(cells, 8)))
    else:  # merged67
        raw = cell(cells, 5)
        segs = raw.split('<br>')
        if len(segs) >= 2:
            rec['mach_parts'].append(segs[0])
            code = '<br>'.join(segs[1:])
        else:
            code = raw
        rec['crew_rows'].append((code, cell(cells, 6), cell(cells, 7)))

def add_continuation(rec, cells, layout):
    c0 = cells[0]
    c0_match = c0.replace('<br>', '')
    if BARE_SUFFIX.match(c0_match):
        rec['shifr_frags'] += [x for x in c0.split('<br>') if x]
    if cell(cells, 1):
        rec['name_parts'].append(cells[1])
    if cell(cells, 2):
        rec['unit_parts'].append(cells[2])
    if cell(cells, 3):
        rec['work_parts'].append(cells[3])
    if cell(cells, 4):
        rec['lab_parts'].append(cells[4])
    if layout == 'std9':
        if cell(cells, 5):
            rec['mach_parts'].append(cells[5])
        if cell(cells, 6) or cell(cells, 7) or cell(cells, 8):
            rec['crew_rows'].append((cell(cells, 6), cell(cells, 7), cell(cells, 8)))
    else:  # merged67
        raw = cell(cells, 5)
        code = ''
        if raw:
            segs = raw.split('<br>')
            if len(segs) >= 2:
                rec['mach_parts'].append(segs[0])
                code = '<br>'.join(segs[1:])
            else:
                code = raw
        if code or cell(cells, 6) or cell(cells, 7):
            rec['crew_rows'].append((code, cell(cells, 6), cell(cells, 7)))

def parse_file(path):
    records = []
    current = None
    layout = 'std9'
    with open(path, encoding='utf-8') as f:
        for lineno, line in enumerate(f, 1):
            cells = split_row(line)
            if cells is None or len(cells) < 3:
                continue
            if is_separator_row(cells):
                continue
            if is_numeric_header_row(cells):
                layout = detect_layout(cells)
                continue
            if is_header_row(cells):
                continue
            if is_noise_row(cells):
                continue
            c0 = cells[0]
            c0_match = c0.replace('<br>', '')
            starts_new = bool(SHIFR_START.match(c0_match))
            if starts_new:
                if current:
                    records.append(current)
                current = new_record(lineno)
                add_start(current, cells, layout)
            else:
                if current is None:
                    continue
                add_continuation(current, cells, layout)
        if current:
            records.append(current)

    out = []
    for r in records:
        shifr = r['shifr_frags'][0] if r['shifr_frags'] else ''
        for frag in r['shifr_frags'][1:]:
            if not shifr.endswith('-') and not frag.startswith('-'):
                shifr += '-'
            shifr += frag
        shifr = shifr.strip()
        name = clean(' '.join(r['name_parts']))
        unit = clean(' '.join(r['unit_parts']))
        lab = clean(' '.join(r['lab_parts']))
        mach = clean(' '.join(r['mach_parts']))

        crew = []
        buf_prof, buf_code = [], []
        for code_cell, prof_cell, qty_cell in r['crew_rows']:
            qty_tokens = [t.strip() for t in qty_cell.split('<br>') if t.strip()]
            prof_tokens = [t.strip() for t in prof_cell.split('<br>') if t.strip()]
            code_tokens = [t.strip() for t in code_cell.split('<br>') if t.strip()]
            if qty_tokens and len(qty_tokens) == len(prof_tokens):
                for i, (pt, qt) in enumerate(zip(prof_tokens, qty_tokens)):
                    if i == 0 and buf_prof:
                        pt = clean(' '.join(buf_prof) + ' ' + pt)
                        buf_prof = []
                    ct = code_tokens[i] if i < len(code_tokens) else (code_tokens[-1] if code_tokens else '')
                    if i == 0 and buf_code and not ct:
                        ct = clean(' '.join(buf_code))
                        buf_code = []
                    crew.append({'code': clean(ct), 'profession': clean(pt), 'qty': qt})
            elif qty_tokens:
                pt = clean(' '.join(buf_prof + [prof_cell]))
                ct = clean(' '.join(buf_code + [code_cell]))
                buf_prof, buf_code = [], []
                crew.append({'code': ct, 'profession': pt, 'qty': qty_tokens[0]})
            else:
                if prof_cell.strip():
                    buf_prof.append(prof_cell)
                if code_cell.strip():
                    buf_code.append(code_cell)

        out.append({
            'line': r['first_line'], 'shifr': shifr, 'name': name, 'unit': unit,
            'lab_hours': lab, 'mach_hours': mach, 'crew': crew,
            'crew_incomplete': bool(buf_prof),
            'valid_shifr': bool(FULL_SHIFR.match(shifr)),
            'source_file': path,
        })
    return out

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--all':
        files = sorted(glob.glob("docs/ФРСН/*.md"))
        total, total_valid, total_no_crew = 0, 0, 0
        for path in files:
            recs = parse_file(path)
            valid = [r for r in recs if r['valid_shifr']]
            no_crew = [r for r in recs if not r['crew']]
            total += len(recs); total_valid += len(valid); total_no_crew += len(no_crew)
            flag = "  <-- проверить" if (len(valid) < len(recs) or no_crew) else ""
            print(f"{len(recs):5d} записей, шифр ok {len(valid):5d}, без состава звена {len(no_crew):4d}  | {path.split('/')[-1][:60]}{flag}")
        print()
        print(f"ИТОГО: {total} записей, шифр корректен {total_valid} ({100*total_valid/total:.1f}%), без состава звена {total_no_crew} ({100*total_no_crew/total:.1f}%)")
    else:
        path = sys.argv[1]
        recs = parse_file(path)
        valid = [r for r in recs if r['valid_shifr']]
        no_crew = [r for r in recs if not r['crew']]
        print(f"Файл: {path}")
        print(f"Записей: {len(recs)}, валидный шифр: {len(valid)}, без состава звена: {len(no_crew)}")
