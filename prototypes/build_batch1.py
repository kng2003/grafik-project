# -*- coding: utf-8 -*-
"""
Шаг 4 фазы 4: первая партия кодов ГЭСН на подтверждение заказчику -
только коды с сильным скором (>=0.7), плюс код, уже решённый вручную
через ЕНиР. Остальные партии (0.5-0.7, 0.3-0.5, и один нерешённый
слабый код) готовятся отдельно после этой.
"""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

SRC = "prototypes/results/2026-09-17_gesn_frsn_candidates_v2.xlsx"
OUT = "prototypes/results/2026-09-19_shag4_partiya1_silnye.xlsx"

wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb["Соответствие ГЭСН-ФРСН"]
rows = list(ws.iter_rows(min_row=2, values_only=True))

def numscore(r):
    v = r[7]
    return v if isinstance(v, (int, float)) else None

strong = [r for r in rows if (numscore(r) is not None and numscore(r) >= 0.7)]
manual_enir = [r for r in rows if r[0] == 'ГЭСНм39-01-002-19']
selected = strong + manual_enir
selected.sort(key=lambda r: (numscore(r) if numscore(r) is not None else 1.0), reverse=True)

out = openpyxl.Workbook()
sh = out.active
sh.title = "Партия 1 - сильные"

headers = ["№", "Код ГЭСН", "Наименование ГЭСН", "Ед.изм.", "Код ФРСН (кандидат)",
           "Наименование ФРСН", "чел-ч ФРСН", "Состав звена", "Скор",
           "Ваше решение (да / нет / уточнить)", "Комментарий"]
sh.append(headers)
for c in sh[1]:
    c.font = Font(bold=True)
    c.alignment = Alignment(wrap_text=True, vertical="top")

yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

for i, r in enumerate(selected, start=1):
    row = [i, r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], "", ""]
    sh.append(row)

widths = [5, 18, 42, 9, 18, 42, 10, 30, 8, 22, 30]
for col, w in zip("ABCDEFGHIJK", widths):
    sh.column_dimensions[col].width = w

# подсветить колонку решения
for row_idx in range(2, sh.max_row + 1):
    sh.cell(row_idx, 10).fill = yellow

sh.freeze_panes = "B2"

sh2 = out.create_sheet("Как заполнить")
instructions = [
    "ЧТО ЭТО ЗА ФАЙЛ",
    "Это первая партия из 4 запланированных. В неё вошли 46 кодов ГЭСН из 157,",
    "у которых алгоритм подобрал кандидата ФРСН с самым высоким скором совпадения",
    "(0.70 и выше по шкале, где 1 - это полное совпадение по смыслу; один код",
    "из этой партии - ГЭСНм39-01-002-19 - подобран не алгоритмом, а вручную",
    "через ЕНиР, отмечен в комментарии).",
    "",
    "ЧТО НУЖНО СДЕЛАТЬ",
    "Для каждой строки посмотреть на 3 колонки: наименование ГЭСН (что делаем),",
    "наименование ФРСН (какая норма подобрана как аналог) и состав звена (кто и",
    "сколько человек нужно по этой норме). Если по смыслу работа та же самая -",
    "написать 'да' в колонке 'Ваше решение'. Если норма явно не подходит -",
    "написать 'нет'. Если не уверены или нужно уточнить у смежников - написать",
    "'уточнить' и коротко пояснить, что смущает, в колонке 'Комментарий'.",
    "",
    "ЧТО ДЕЛАТЬ С ОТВЕТОМ 'НЕТ'",
    "Ничего страшного - у каждого кода в полном файле",
    "2026-09-17_gesn_frsn_candidates_v2.xlsx есть ещё 2 кандидата на выбор.",
    "Просто напишите 'нет' и в комментарии, если знаете, что здесь должно быть",
    "на самом деле - этого достаточно, дальше разберём вместе.",
    "",
    "СКОЛЬКО ВРЕМЕНИ ЭТО ЗАЙМЁТ",
    "46 строк, большинство - на 10-15 секунд каждая, потому что скор высокий",
    "и совпадение обычно очевидное на глаз. Прикидочно 15-25 минут на всю партию.",
    "",
    "ЧТО ДАЛЬШЕ",
    "После этой партии подготовлю вторую - коды со скором 0.5-0.7 (65 штук),",
    "затем третью - 0.3-0.5 (45 штук), и отдельно - 2 кода, где алгоритм не",
    "справился совсем (один уже решён через ЕНиР, второй пока без ответа).",
]
for line in instructions:
    sh2.append([line])
sh2.column_dimensions['A'].width = 95
for row_idx in range(1, len(instructions) + 1):
    sh2.cell(row_idx, 1).alignment = Alignment(wrap_text=True)
sh2['A1'].font = Font(bold=True, size=13)

out.save(OUT)
print(f"Сохранено: {OUT}")
print(f"Строк в партии 1: {len(selected)}")
