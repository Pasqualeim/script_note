import os
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill

# === Sezione 1: Confronto e Colorazione delle Note Impattate ===
def apply_color_to_note_number(components_df, notes_df, output_filename="Note Extraction_Updated.xlsx", red_notes_filename="Impacted_Notes.xlsx"):
    if not os.path.exists("Note Extraction.xlsx"):
        print("❌ Errore: Il file Note Extraction.xlsx non esiste.")
        return

    wb = load_workbook("Note Extraction.xlsx")
    ws = wb.active  

    # Creazione di un nuovo file Excel per le sole note impattate
    wb_red_notes = Workbook()
    ws_red_notes = wb_red_notes.active
    ws_red_notes.append(["Note Number", "Impacted Component", "Software Component", "From", "To", "Patch Level (Note Extraction)", "SPLevel (Component)"])

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for index, note_row in notes_df.iterrows():
        found = False
        impacted_components = []
        software_component = note_row.get('Software Component', None)
        patch_level_note = note_row.get('Patch Level', None)
        sp_level_component = None
        note_number = note_row.get('Note Number')

        for _, component_row in components_df.iterrows():
            sp_level = component_row.get('SPLevel', None)

            if patch_level_note is None or (sp_level is not None and isinstance(patch_level_note, (int, float)) and patch_level_note > sp_level):
                found = True
                impacted_components.append(component_row['Component'])
                sp_level_component = sp_level

        if found:
            note_row_idx = index + 2  
            note_cell = ws[f"A{note_row_idx}"]  
            note_cell.fill = red_fill
            print(f"🔴 Impattato: {note_cell.coordinate} colorato di rosso")

            ws_red_notes.append([
                note_cell.value,
                ", ".join(set(impacted_components)),
                software_component,
                note_row.get('From'),
                note_row.get('To'),
                patch_level_note,
                sp_level_component
            ])

    wb.save(output_filename)
    wb_red_notes.save(red_notes_filename)
    print(f"✅ File aggiornato e salvato: {output_filename} e {red_notes_filename}")

# === Sezione 2: Eliminazione delle Note Duplicate ===
def delete_duplicate_notes(file_path="Impacted_Notes.xlsx"):
    if not os.path.exists(file_path):
        print("❌ Errore: Il file Impacted_Notes.xlsx non esiste.")
        return

    wb = load_workbook(file_path)
    ws = wb.active

    # Trova la colonna della "Note Number"
    note_number_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == "Note Number":
            note_number_col = col
            break

    if note_number_col:
        previous_note = None
        for row in range(2, ws.max_row + 1):
            current_note = ws.cell(row=row, column=note_number_col).value

            if current_note == previous_note:
                ws.cell(row=row, column=note_number_col, value="")  # Svuota la cella duplicata
            else:
                previous_note = current_note

    updated_file_path = "Impacted_Notes_Cleaned.xlsx"
    wb.save(updated_file_path)
    print(f"✅ File pulito salvato come: {updated_file_path}")

# === Sezione 3: Unione Celle delle Note Duplicate ===
def merge_duplicate_notes(file_path="Impacted_Notes_Cleaned.xlsx"):
    if not os.path.exists(file_path):
        print("❌ Errore: Il file Impacted_Notes_Cleaned.xlsx non esiste.")
        return

    wb = load_workbook(file_path)
    ws = wb.active

    # Trova la colonna della "Note Number"
    note_number_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == "Note Number":
            note_number_col = col
            break

    if note_number_col:
        previous_note = None
        start_row = None
        note_values = []

        for row in range(2, ws.max_row + 1):
            current_note = ws.cell(row=row, column=note_number_col).value

            if current_note and current_note == previous_note:
                if start_row is None:
                    start_row = row - 1
                note_values.append(current_note)
            else:
                if start_row is not None and start_row < row - 1:
                    merged_value = ", ".join(set(note_values))
                    ws.cell(row=start_row + 1, column=note_number_col, value=merged_value)
                    ws.merge_cells(start_row=start_row + 1, start_column=note_number_col, end_row=row - 1, end_column=note_number_col)
                start_row = None
                note_values = [current_note] if current_note else []

            previous_note = current_note

        if start_row is not None and start_row < ws.max_row:
            merged_value = ", ".join(set(note_values))
            ws.cell(row=start_row + 1, column=note_number_col, value=merged_value)
            ws.merge_cells(start_row=start_row + 1, start_column=note_number_col, end_row=ws.max_row, end_column=note_number_col)

    updated_file_path = "Impacted_Notes_Merged.xlsx"
    wb.save(updated_file_path)
    print(f"✅ File con celle unite salvato come: {updated_file_path}")

# === Esecuzione delle Funzioni ===
if __name__ == "__main__":
    components_df = pd.read_excel("Components.xlsx")
    notes_df = pd.read_excel("Note Extraction.xlsx")

    apply_color_to_note_number(components_df, notes_df)
    delete_duplicate_notes()
    merge_duplicate_notes()