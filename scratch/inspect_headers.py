import openpyxl
import glob
import os
import sys

# Ensure UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

for f in sorted(glob.glob('đầu vào/*/*.xlsx')):
    wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
    ws = wb.active
    rows = []
    for r in range(1, 15):
        row_vals = [ws.cell(r, c).value for c in range(1, 10)]
        if any(row_vals):
            rows.append((r, [v for v in row_vals if v is not None]))
    print(f"=== FILE: {f}")
    for r_num, vals in rows[:7]:
        print(f"  Row {r_num}: {vals[:5]}")
