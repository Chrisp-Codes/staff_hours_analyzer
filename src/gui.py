import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import os
import sys
from openpyxl import load_workbook
from openpyxl.styles import Alignment

def calculate_hourly_presence(file_path):
    df = pd.read_excel(file_path, sheet_name="Alle Mitarbeiter", skiprows=6)

    # Nur relevante Spalten
    df = df[["Tag", "Startzeit", "Endzeit", "Dauer netto (dezimal)"]].dropna()

    # Tag in echtes Datum umwandeln (robust)
    df = df[df["Tag"].notna()]
    df["Datum"] = df["Tag"].apply(lambda x: x.split()[1] if isinstance(x, str) and len(x.split()) > 1 else None)
    df = df[df["Datum"].notna()]
    df["Datum"] = pd.to_datetime(df["Datum"], format="%d.%m.%Y")

    # Zeitkonvertierung
    df["Start"] = pd.to_datetime(df["Datum"].astype(str) + " " + df["Startzeit"].astype(str))
    df["Ende"] = pd.to_datetime(df["Datum"].astype(str) + " " + df["Endzeit"].astype(str))

    def split_row(row):
        start = row["Start"]
        end = row["Ende"]
        dauer = row["Dauer netto (dezimal)"]
        total_minutes = dauer * 60
        result = []

        current = start
        while current < end:
            next_hour = (current + timedelta(hours=1)).replace(minute=0, second=0)
            if next_hour > end:
                next_hour = end

            minutes_in_this_hour = (next_hour - current).total_seconds() / 60
            ratio = minutes_in_this_hour / total_minutes if total_minutes > 0 else 0

            result.append({
                "Datum": row["Datum"].date(),
                "Stunde": current.strftime("%H:00"),
                "Personalstunden": dauer * ratio
            })
            current = next_hour

        return result

    rows = []
    for _, row in df.iterrows():
        rows.extend(split_row(row))

    df_expanded = pd.DataFrame(rows)
    df_grouped = df_expanded.groupby(["Datum", "Stunde"]).sum().reset_index()
    df_grouped["Personalstunden"] = df_grouped["Personalstunden"].round(2)

    # Formatierter Export
    output_rows = []
    for datum in df_grouped["Datum"].unique():
        output_rows.append([""])
        output_rows.append([datum.strftime("%Y-%m-%d")])
        for _, row in df_grouped[df_grouped["Datum"] == datum].iterrows():
            output_rows.append(["", row["Stunde"], row["Personalstunden"]])

    result_df = pd.DataFrame(output_rows)

    original_name = os.path.basename(file_path)
    base_name = os.path.splitext(original_name)[0]
    export_path = os.path.join(os.path.dirname(file_path), f"Stundenanalyse_{base_name}.xlsx")
    result_df.to_excel(export_path, index=False, header=["Datum", "Stunde", "Personalstunden"])

    # Formatierung anwenden (openpyxl)
    wb = load_workbook(export_path)
    ws = wb.active
    for col in ["B", "C"]:
        ws.column_dimensions[col].width = 20
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=3):
        for cell in row:
            cell.alignment = Alignment(horizontal="center")
    wb.save(export_path)

    messagebox.showinfo("Fertig", f"Auswertung gespeichert: {export_path}")
    sys.exit()

def main():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Bitte Export auswählen", filetypes=[("Excel-Dateien", "*.xlsx")])
    if file_path:
        calculate_hourly_presence(file_path)

if __name__ == "__main__":
    main()
