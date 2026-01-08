import pandas as pd
import sys
import os
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils.dataframe import dataframe_to_rows

def get_formula_templates(sheet, header_row=1):
    formulas = {}
    data_row_idx = header_row + 1
    
    headers = [cell.value for cell in sheet[header_row]]
    
    if sheet.max_row < data_row_idx:
        return {}

    for col_idx, cell in enumerate(sheet[data_row_idx], start=0):
        if isinstance(cell.value, str) and cell.value.startswith("="):
            col_name = headers[col_idx] if col_idx < len(headers) else None
            if col_name:
                formulas[col_name] = cell.value
                print(f"-> Formel erkannt in Spalte '{col_name}': {cell.value}")
    
    return formulas

def main():
    src_path = input("Pfad zur Quelldatei (.xlsx): ").strip().strip('"')
    tgt_path = input("Pfad zur Zieldatei (.xlsx): ").strip().strip('"')

    if not os.path.exists(src_path) or not os.path.exists(tgt_path):
        print("Fehler: Dateien nicht gefunden.")
        sys.exit(1)

    print("Lade Dateien...")
    try:
        wb = load_workbook(tgt_path)
        ws = wb.active
        formula_map = get_formula_templates(ws)
        wb.close()

        df_src = pd.read_excel(src_path)
        df_tgt = pd.read_excel(tgt_path)
        
        df_src.columns = df_src.columns.astype(str)
        df_tgt.columns = df_tgt.columns.astype(str)

    except Exception as e:
        print(f"Lese-Fehler: {e}")
        sys.exit(1)

    print("\n--- Spalten-Zuordnung ---")
    tgt_cols = df_tgt.columns.tolist()
    src_cols = df_src.columns.tolist()
    mapping = {}

    print("Verfügbare Quell-Spalten:")
    for col in src_cols:
        print(f" - {col}")

    print("\nVerfügbare Ziel-Spalten:")
    for col in tgt_cols:
        print(f" - {col}")

    for tgt_col in tgt_cols:
        is_formula = tgt_col in formula_map
        note = " (FORMEL)" if is_formula else ""
        
        if is_formula:
            print(f"Ziel '{tgt_col}' wird automatisch berechnet.{note}")
            continue

        default = tgt_col if tgt_col in src_cols else ""
        suggestion = f" [Enter für '{default}']" if default else ""

        user_input = input(f"Ziel '{tgt_col}' <- Quelle?{suggestion}: ").strip()
        src_col = user_input if user_input else default
        
        if src_col and src_col in src_cols:
            mapping[tgt_col] = src_col

    data_to_append = pd.DataFrame()
    for tgt_col, src_col in mapping.items():
        data_to_append[tgt_col] = df_src[src_col]

    df_combined = pd.concat([df_tgt, data_to_append], ignore_index=True)

    print("\n--- Sortierung ---")
    prio1 = input("Priorität 1 Spalte (Name): ").strip()
    prio2 = input("Priorität 2 Spalte (Name): ").strip()

    sort_cols = []
    for col_name in [prio1, prio2]:
        if col_name and col_name in df_combined.columns:
            if pd.api.types.is_datetime64_any_dtype(df_combined[col_name]):
                sort_cols.append(col_name)
                continue

            try:
                df_combined[col_name] = pd.to_datetime(
                    df_combined[col_name], 
                    dayfirst=True, 
                    errors='coerce'
                )
                sort_cols.append(col_name)
            except Exception:
                sort_cols.append(col_name)

    if sort_cols:
        df_combined = df_combined.sort_values(
            by=sort_cols, 
            ascending=[True]*len(sort_cols),
            na_position='last'
        )
        print(f"Sortiert nach {sort_cols}")

    print(f"\nSchreibe Datei: {tgt_path} ...")
    
    wb = load_workbook(tgt_path)
    ws = wb.active
    
    if ws.max_row > 1:
        ws.delete_rows(2, amount=ws.max_row-1)

    rows = dataframe_to_rows(df_combined, index=False, header=False)
    
    header_list = list(df_combined.columns)
    
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            col_name = header_list[c_idx-1]
            cell = ws.cell(row=r_idx, column=c_idx)

            if col_name in formula_map:
                original_formula = formula_map[col_name]
                
                try:
                    translated = Translator(original_formula, origin=f"A2").translate_formula(f"A{r_idx}")
                    cell.value = translated
                except Exception:
                    cell.value = value 
            else:
                if pd.isna(value):
                    cell.value = None
                else:
                    cell.value = value

    wb.save(tgt_path)
    print("Fertig! Formeln wurden auf alle Zeilen angewendet.")

if __name__ == "__main__":
    main()