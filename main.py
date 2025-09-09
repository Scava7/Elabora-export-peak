import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import pandas as pd

def parse_trc(trc_text: str):
    lines = trc_text.splitlines()
    info_lines = lines[:6]  # first 6 lines "as-is"
    records = []
    for ln in lines[6:]:
        if ln.strip().startswith(";"):
            continue
        if ")" in ln:
            try:
                clean = ln.strip().replace(")", "")
                parts = clean.split()
                if len(parts) >= 5 and parts[0].isdigit():
                    msg_no = int(parts[0])
                    time_offset = float(parts[1].replace(",", "."))
                    msg_type = parts[2]
                    id_hex = parts[3].upper()
                    dlc = int(parts[4])
                    data_bytes = [b.upper() for b in parts[5:5+8]]
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
                pass
    records.sort(key=lambda r: r["Message #"])
    return info_lines, records

def export_xlsx(in_path: Path) -> Path:
    text = in_path.read_text(errors="replace")
    info_lines, records = parse_trc(text)
    df = pd.DataFrame(records, columns=[
        "Message #", "Time Offset (ms)", "Type", "ID (hex)", "DLC",
        "Byte0", "Byte1", "Byte2", "Byte3", "Byte4", "Byte5", "Byte6", "Byte7"
    ])

    out_path = in_path.with_suffix(".xlsx")
    # Write using ExcelWriter so we can place the table after the first 6 lines
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Write the table starting from row 8 (index 7 => 0-based)
        start_row = 7  # after 6 info lines + one blank row
        df.to_excel(writer, sheet_name="Trace", index=False, startrow=start_row)
        ws = writer.sheets["Trace"]
        # Write the first 6 lines at the top
        for i, line in enumerate(info_lines, start=1):
            ws.cell(row=i, column=1, value=line)

        # Optional: set column widths for readability
        widths = {
            "A": 12, "B": 16, "C": 10, "D": 12, "E": 8,
            "F": 8, "G": 8, "H": 8, "I": 8, "J": 8, "K": 8, "L": 8, "M": 8,
        }
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    return out_path

def main():
    root = tk.Tk()
    root.withdraw()  # hide main window
    path = filedialog.askopenfilename(
        title="Seleziona il file PCAN Trace (.trc)",
        filetypes=[("PCAN Trace", "*.trc"), ("Testo", "*.txt"), ("Tutti i file", "*.*")]
    )
    if not path:
        return

    try:
        out_path = export_xlsx(Path(path))
        messagebox.showinfo("Operazione completata", f"Excel creato:\n{out_path}")
    except Exception as e:
        messagebox.showerror("Errore", f"Impossibile creare l'Excel:\n{e}")

if __name__ == "__main__":
    main()