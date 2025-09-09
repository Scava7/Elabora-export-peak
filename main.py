# -*- coding: utf-8 -*-
"""
PCAN .trc -> Excel .xlsx (tabulato) con evidenziazione per 3 ID impostabili in G2, G3, G4.
- Prime 6 righe del file copiate "as-is" in alto.
- Header tabella alla riga 8, dati dalla riga 9 in giù.
- Colori: G2 rosso chiaro, G3 verde chiaro, G4 azzurro (match "contains" sull'ID in colonna D).
"""

import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule


def parse_trc(trc_text: str):
    """Parsa un PCAN-View .trc (v5) in (prime 6 righe info, records tabella)."""
    lines = trc_text.splitlines()
    info_lines = lines[:6]  # mantieni le prime 6 righe così come sono

    records = []
    for ln in lines[6:]:
        if ln.strip().startswith(";"):
            continue  # ignora righe commento
        # Esempio riga messaggio:
        # "     2)         0.8  Rx         0086  8  00 80 15 00 00 00 00 00"
        if ")" in ln:
            try:
                clean = ln.strip().replace(")", "")
                parts = clean.split()
                # parts atteso: [msg_no, time_offset, type, id_hex, dlc, byte0..byte7]
                if len(parts) >= 5 and parts[0].isdigit():
                    msg_no = int(parts[0])
                    # gestisci eventuale virgola decimale
                    time_offset = float(parts[1].replace(",", "."))
                    msg_type = parts[2]
                    id_hex = parts[3].upper()
                    dlc = int(parts[4])
                    data_bytes = [b.upper() for b in parts[5:5 + 8]]
                    # pad a 8 byte
                    data_bytes += [""] * (8 - len(data_bytes))

                    records.append({
                        "Message #": msg_no,
                        "Time Offset (ms)": time_offset,
                        "Type": msg_type,
                        "ID (hex)": id_hex,
                        "DLC": dlc,
                        "Byte0": data_bytes[0],
                        "Byte1": data_bytes[1],
                        "Byte2": data_bytes[2],
                        "Byte3": data_bytes[3],
                        "Byte4": data_bytes[4],
                        "Byte5": data_bytes[5],
                        "Byte6": data_bytes[6],
                        "Byte7": data_bytes[7],
                    })
            except Exception:
                # ignora righe malformate
                pass

    records.sort(key=lambda r: r["Message #"])
    return info_lines, records


def add_id_highlights(ws, first_data_row: int, last_data_row: int):
    """Aggiunge 3 regole di formattazione condizionale basate su G2, G3, G4."""
    # Etichette di aiuto
    ws["F2"] = "ID filtro #1"
    ws["F3"] = "ID filtro #2"
    ws["F4"] = "ID filtro #3"

    # Riempimenti (ARGB a 8 cifre, con alpha "FF")
    red_fill = PatternFill(start_color="FFF4CCCC", end_color="FFF4CCCC", fill_type="solid")
    green_fill = PatternFill(start_color="FFD9EAD3", end_color="FFD9EAD3", fill_type="solid")
    blue_fill = PatternFill(start_color="FFCCFFFF", end_color="FFCCFFFF", fill_type="solid")

    # Applica su tutto l'intervallo dati (A..M)
    data_range = f"A{first_data_row}:M{last_data_row}"

    # La formula fa riferimento alla cella ID della PRIMA riga del range (D<first_row>).
    # Excel adatterà la riga automaticamente per ogni riga del range.
    f1 = f'=AND($G$2<>"",NOT(ISERROR(SEARCH($G$2,$D{first_data_row}))))'
    f2 = f'=AND($G$3<>"",NOT(ISERROR(SEARCH($G$3,$D{first_data_row}))))'
    f3 = f'=AND($G$4<>"",NOT(ISERROR(SEARCH($G$4,$D{first_data_row}))))'

    ws.conditional_formatting.add(data_range, FormulaRule(formula=[f1], fill=red_fill))
    ws.conditional_formatting.add(data_range, FormulaRule(formula=[f2], fill=green_fill))
    ws.conditional_formatting.add(data_range, FormulaRule(formula=[f3], fill=blue_fill))

    # Blocca riga intestazione per lo scroll
    ws.freeze_panes = f"A{first_data_row}"


def export_xlsx(in_path: Path) -> Path:
    """Legge .trc e scrive .xlsx con tabella e formattazione condizionale per gli ID."""
    text = in_path.read_text(errors="replace")
    info_lines, records = parse_trc(text)

    df = pd.DataFrame(records, columns=[
        "Message #", "Time Offset (ms)", "Type", "ID (hex)", "DLC",
        "Byte0", "Byte1", "Byte2", "Byte3", "Byte4", "Byte5", "Byte6", "Byte7"
    ])

    out_path = in_path.with_suffix(".xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        start_row = 7  # -> header a riga 8, dati da riga 9
        df.to_excel(writer, sheet_name="Trace", index=False, startrow=start_row)
        ws = writer.sheets["Trace"]

        # Scrivi le prime 6 righe "as-is"
        for i, line in enumerate(info_lines, start=1):
            ws.cell(row=i, column=1, value=line)

        # Larghezze colonne
        widths = {
            "A": 12, "B": 16, "C": 10, "D": 12, "E": 8,
            "F": 8, "G": 12, "H": 8, "I": 8, "J": 8, "K": 8, "L": 8, "M": 8,
        }
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

        # AutoFilter su tutto il range tabellare
        first_data_row = start_row + 2      # 9
        last_data_row = first_data_row + len(df) - 1 if len(df) else first_data_row
        ws.auto_filter.ref = f"A{start_row+1}:M{last_data_row}"

        # Formattazione condizionale basata su G2,G3,G4
        add_id_highlights(ws, first_data_row, last_data_row)

    return out_path


def main():
    root = tk.Tk()
    root.withdraw()  # niente finestra principale

    path = filedialog.askopenfilename(
        title="Seleziona il file PCAN Trace (.trc)",
        filetypes=[("PCAN Trace", "*.trc"), ("Testo", "*.txt"), ("Tutti i file", "*.*")]
    )
    if not path:
        return

    try:
        out = export_xlsx(Path(path))
        messagebox.showinfo(
            "Operazione completata",
            f"Excel creato:\n{out}\n\nInserisci gli ID in G2, G3, G4 per evidenziare le righe."
        )
    except Exception as e:
        messagebox.showerror("Errore", f"Impossibile creare l'Excel:\n{e}")


if __name__ == "__main__":
    main()
