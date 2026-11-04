# -*- coding: utf-8 -*-
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('Receptores_panaderia.xlsx', data_only=True)
for sheet_name in wb.sheetnames:
    sheet = wb[sheet_name]
    print(f"=== Sheet: {sheet_name} (rows: {sheet.max_row}, cols: {sheet.max_column}) ===")
    for r in range(1, sheet.max_row + 1):
        vals = [sheet.cell(r, c).value for c in range(1, sheet.max_column + 1)]
        if any(v is not None for v in vals):
            clean_vals = [str(v).strip() if v is not None else "" for v in vals]
            while clean_vals and not clean_vals[-1]:
                clean_vals.pop()
            print(f"L{r:02d}: " + " | ".join(clean_vals))
