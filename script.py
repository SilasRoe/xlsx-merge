import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from copy import copy
import os
from datetime import datetime, timedelta

def to_excel_serial(val):
    if pd.isna(val): return None
    delta = val - datetime(1899, 12, 30)
    return float(delta.days) + (float(delta.seconds) / 86400)

def from_excel_serial(val):
    if pd.isna(val) or not isinstance(val, (int, float)): return None
    return datetime(1899, 12, 30) + timedelta(days=val)

def merge_and_sort_excel(source_path: str, target_path: str):
    if not os.path.exists(source_path) or not os.path.exists(target_path):
        print("Fehler: Dateien nicht gefunden.")
        return

    print("Lade Zieldatei...")
    wb = load_workbook(target_path)
    ws = wb.active
    
    target_headers = {cell.value: cell.column for cell in ws[1] if cell.value}
    header_names = list(target_headers.keys())

    template_info = {}
    if ws.max_row >= 2:
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=2, column=col_idx)
            template_info[col_idx] = {
                'font': copy(cell.font), 'border': copy(cell.border),
                'fill': copy(cell.fill), 'number_format': copy(cell.number_format),
                'alignment': copy(cell.alignment), 'protection': copy(cell.protection),
                'is_formula': cell.data_type == 'f',
                'formula': cell.value if cell.data_type == 'f' else None,
                'col_letter': cell.column_letter
            }

    df_target_old = pd.read_excel(target_path)
    df_target_old = df_target_old.loc[:, ~df_target_old.columns.str.contains('^Unnamed')]
    df_target_old = df_target_old[header_names] if not df_target_old.empty else pd.DataFrame(columns=header_names)

    print(f"Lade Quelle '{source_path}'...")
    df_source = pd.read_excel(source_path)
    
    print("\n--- Spaltenzuordnung ---")
    print("Verfügbare Quell-Spalten:" + f" {', '.join(df_source.columns)})")
    print("Verfügbare Ziel-Spalten:" + f" {', '.join(df_target_old.columns)})")
    mapping = {}
    for t_col in header_names:
        col_idx = target_headers[t_col]
        if template_info.get(col_idx, {}).get('is_formula'): continue
        
        s_col = input(f"Quelle für '{t_col}'? (Enter=leer): ").strip()
        if s_col in df_source.columns: mapping[t_col] = s_col

    df_new_data = pd.DataFrame()
    for t_col in header_names:
        df_new_data[t_col] = df_source[mapping[t_col]] if t_col in mapping else None

    df_total = pd.concat([df_target_old, df_new_data], ignore_index=True)

    date_cols = []
    print("Bereite Daten für Sortierung vor...")
    for col in df_total.columns:
        if pd.api.types.is_datetime64_any_dtype(df_total[col]):
            date_cols.append(col)
            df_total[col] = df_total[col].apply(lambda x: to_excel_serial(pd.to_datetime(x)) if pd.notnull(x) else None)

    print("\n--- Sortierung ---")
    sort_col = input(f"Sortieren nach? ({', '.join(header_names)}): ").strip()
    
    if sort_col in df_total.columns:
        print(f"Sortiere nach '{sort_col}'...")
        try:
            df_total = df_total.sort_values(by=sort_col)
        except TypeError:
            print(f"Warnung: Gemischte Typen in '{sort_col}'. Sortiere als Text.")
            df_total = df_total.sort_values(by=sort_col, key=lambda x: x.astype(str))

    if date_cols:
        print(f"Wandle Datumsspalten zurück: {date_cols}")
        for col in date_cols:
            df_total[col] = df_total[col].apply(lambda x: from_excel_serial(x))

    print("Schreibe Daten...")
    start_row = 2
    for i, (_, row_data) in enumerate(df_total.iterrows()):
        current_row = start_row + i
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=current_row, column=col_idx)
            t_info = template_info.get(col_idx)
            if not t_info: continue

            if t_info['is_formula'] and t_info['formula']:
                orig = f"{t_info['col_letter']}2"
                targ = f"{t_info['col_letter']}{current_row}"
                cell.value = Translator(t_info['formula'], origin=orig).translate_formula(targ)
            else:
                header_name = ws.cell(row=1, column=col_idx).value
                val = row_data.get(header_name)
                cell.value = val if (pd.notnull(val) and val is not None) else None

            if t_info['number_format']: cell.number_format = copy(t_info['number_format'])
            cell.font = copy(t_info['font'])
            cell.border = copy(t_info['border'])
            cell.fill = copy(t_info['fill'])
            cell.alignment = copy(t_info['alignment'])
            cell.protection = copy(t_info['protection'])

    try:
        wb.save(target_path)
        print(f"Fertig! ({len(df_total)} Zeilen)")
    except PermissionError:
        print("Fehler: Datei ist noch geöffnet!")

if __name__ == "__main__":
    s = input("Quell-Datei: ").strip().strip('"')
    t = input("Ziel-Datei: ").strip().strip('"')
    merge_and_sort_excel(s, t)