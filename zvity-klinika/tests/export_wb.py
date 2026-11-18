"""Вивантажує аркуші шаблону в JSON для tests/test_kod_mock.js.  python3 tests/export_wb.py Zvity_kliniki.xlsx out.json"""
import json, sys, datetime
from openpyxl import load_workbook
wb = load_workbook(sys.argv[1])
out = {}
for ws in wb.worksheets:
    if ws.title in ('Дашборд', 'Рік', 'Інструкція'): continue
    out[ws.title] = [[{'$d': v.isoformat()} if isinstance(v, (datetime.date, datetime.datetime)) else ('' if v is None else v) for v in row]
                     for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 3001), max_col=min(ws.max_column, 26), values_only=True)]
json.dump(out, open(sys.argv[2], 'w'), ensure_ascii=False)
