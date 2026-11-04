# -*- coding: utf-8 -*-
import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

# Localizar archivo tanto si se ejecuta desde raíz como desde 'ajustes gemini'
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
excel_path = os.path.join(base_dir, 'calculos', 'Receptores_panaderia_v2.xlsx')

wb = openpyxl.load_workbook(excel_path, data_only=True)
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
