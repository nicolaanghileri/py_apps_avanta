# logic.py

import os
import openpyxl

# Globale Dokumentnummer, wird in pivot_kst_columns hochgezählt
doc_number = 1

def pivot_kst_columns(source_file, prefix):

    global doc_number

    wb_source = openpyxl.load_workbook(source_file, data_only=True)
    ws_source = wb_source.active

    # Header-Zeile im Quell-Sheet
    headers = [cell.value for cell in ws_source[1]]

    # Relevante Spalten anhand der Headers ermitteln
    try:
        col_bez1 = headers.index("Bezeichnung 1")
        col_kto2 = headers.index("Kontonummer 2")
        col_bez2 = headers.index("Bezeichnung 2")
    except ValueError as e:
        print(f"Fehlender Header in {source_file}: {e}")
        wb_source.close()
        return []

    # Alle "KST"-Spalten finden
    kst_cols = []
    for i, h in enumerate(headers):
        if h and str(h).startswith("KST"):
            kst_cols.append(i)

    rows_to_append = []
      # Startwert für die Dokumentnummer

    # Zeilen durchlaufen (ab Zeile 5)
    for row in ws_source.iter_rows(min_row=5, values_only=True):
        bezeichnung1 = row[col_bez1]
        kontonummer2 = row[col_kto2]
        bezeichnung2 = row[col_bez2]

        for c_idx in kst_cols:
            amount = row[c_idx]
            if amount and amount != 0:
                # Extrahiere KST-Nummer aus dem Header
                cost_center_num = headers[c_idx].split()[1]
                # Kombiniere prefix und KST-Nummer
                new_kst = prefix + cost_center_num

                # HIER bauen wir direkt die 12 Spalten in der korrekten Reihenfolge
                rows_to_append.append([
                    "01.02.2025",   # Value Date (Hardcoded)
                    "-",            # Buchungsart
                    prefix,         # Company number
                    new_kst,        # Cost Center
                    kontonummer2,   # Cost type
                    bezeichnung2,   # Cost type description
                    "CHF",          # Currency
                    float(amount),         # Amount
                    "01.02.2025",   # Entry date (hier beispielhaft gleiches Datum)
                    "",             # Local Fibu Account (leer oder anpassen)
                    doc_number,     # Document Number
                    ""              # Document Description (leer)
                ])

                doc_number += 1

    wb_source.close()
    return rows_to_append


def checksum(source_file, min_row, column):
    """
    Summiert alle numerischen Werte (float/int) in einer bestimmten Spalte
    eines Excel-Files ab einer definierten Zeile.
    """
    wb_source = openpyxl.load_workbook(source_file, data_only=True)
    ws_source = wb_source.active

    sum_value = 0.0
    for row in ws_source.iter_rows(min_row=min_row, values_only=True):
        value = row[column]
        if value is None:
            continue
        try:
            value_as_float = float(value)
            sum_value += value_as_float
        except ValueError:
            print(f"Warnung: '{value}' in Spalte {column} ist nicht numerisch und wird übersprungen. "
                  f"Source: {source_file}")

    wb_source.close()
    return sum_value


def floats_equal(a, b, epsilon=1e-9):
    """
    Nützliche Hilfsfunktion, um zwei Fließkommazahlen tolerant zu vergleichen.
    """
    return abs(a - b) < epsilon


def main_algo(source_dir, output_dir, label_status=None):
    """
    Diese Funktion durchläuft alle Excel-Dateien im Quellverzeichnis (source_dir),
    holt sich die relevanten Daten (mittels pivot_kst_columns),
    führt eine Checksum-Prüfung durch und schreibt ein konsolidiertes
    Ergebnis in eine neue Datei 'MonthlyReport_STAG.xlsx' im Zielverzeichnis (output_dir).

    label_status ist optional und dient zur Ausgabe von Statusmeldungen in der GUI.
    """

    # Liste aller Dateien im Ordner (Filtern nach .xlsx)
    files = [
        f for f in os.listdir(source_dir)
        if os.path.isfile(os.path.join(source_dir, f)) and f.endswith(".xlsx")
    ]

    final_rows = []
    final_checksum_files = 0

    # Durchlaufe alle gefundenen Dateien
    for f in files:
        file_path = os.path.join(source_dir, f)
        prefix = f.split()[0]  # Nimmt den ersten Wert aus dem Dateinamen als Prefix/Hotelnummer

        # Addiere die Checksum über die Spalte 4 (Index 4 => 5. Spalte) ab Zeile 5
        final_checksum_files += checksum(file_path, 5, 4)

        rows = pivot_kst_columns(file_path, prefix)
        print(f"Verarbeite Datei: {file_path} mit prefix='{prefix}'")
        if rows:
            final_rows.extend(rows)

    print(f"Anzahl gesammelter Zeilen: {len(final_rows)}")

    # Neues Workbook erstellen und Daten in "Pivot_Gesamt" schreiben
    wb_pivot = openpyxl.Workbook()
    ws_pivot = wb_pivot.active
    ws_pivot.title = "Pivot_Gesamt"

    # Header für die 12 Spalten
    header = [
        "Value Date",
        "Buchungsart",
        "Company number",
        "Cost Center",
        "Cost type",
        "Cost type description",
        "Currency",
        "Amount",
        "Entry date",
        "Local Fibu Account",
        "Document Number",
        "Document Description"
    ]
    ws_pivot.append(header)

    # Alle gesammelten Zeilen einfügen
    for row in final_rows:
        ws_pivot.append(row)

    # Speichern
    output_xlsx = os.path.join(output_dir, "MonthlyReport_STAG.xlsx")
    wb_pivot.save(output_xlsx)

    # Checksum Kontrolle
    final_checksum_total_amount = checksum(output_xlsx, 2, 7)

    print(final_checksum_files)
    print(final_checksum_total_amount)
    if floats_equal(final_checksum_files, final_checksum_total_amount):
        print("CHECKSUM AMOUNT -> OK")
        print(final_checksum_files)
        print(final_checksum_total_amount)
        if label_status:
            label_status.config(text="EXPORT SUCCESS", fg="green")
    else:
        print("CHECKSUM AMOUNT -> NOT OK!!")
        if label_status:
            label_status.config(text="EXPORT FAILED", fg="red")

    print(f"Fertig! Alle Daten in '{output_xlsx}' gespeichert.")
