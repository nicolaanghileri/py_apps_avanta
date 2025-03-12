import os
import glob
import openpyxl

first_map = []

def load_first_file(file_path):
    """ Erstellt die Basisstruktur für first_map mit allen Kontonummern. """
    global first_map
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    for row in ws.iter_rows(values_only=True, min_row=2):
        first_map.append([row[0], row[1], None, None, None, None])  # Basis erstellen


def populate_totals(file_path, pos):
    """ Füllt die Werte für die jeweilige Datei in die entsprechende Spalte ein. """
    global first_map
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    konto_dict = {konto[0]: konto for konto in first_map}  # Schnellzugriff für Matching

    for row in ws.iter_rows(values_only=True, min_row=2):
        konto_nr = row[0]
        saldo = row[3]  # Annahme: Spalte 4 enthält den Saldo

        if konto_nr in konto_dict:
            konto_dict[konto_nr][pos] = saldo  # Setze den Wert in first_map
        else:
            print(f"Warnung: Konto {konto_nr} in {file_path} nicht gefunden!")


def process_files(source_dir):
    """ Durchsucht den Ordner nach .xlsx Dateien und verarbeitet sie. """
    files = glob.glob(os.path.join(source_dir, "*.xlsx"))

    if not files:
        print("Keine Dateien gefunden!")
        return

    # Erste Datei als Basis nehmen
    load_first_file(files[0])

    # Restliche Dateien nach Präfix bearbeiten
    for file_path in files:
        filename = os.path.basename(file_path)
        prefix = filename.split()[0]  # Annahme: Präfix ist die erste Zahl im Dateinamen

        if prefix == "061":
            populate_totals(file_path, 2)
        elif prefix == "486":
            populate_totals(file_path, 3)
        elif prefix == "495":
            populate_totals(file_path, 4)
        elif prefix == "725":
            populate_totals(file_path, 5)
        else:
            print(f"Unbekanntes Präfix in Datei {filename}, wird übersprungen.")

    # Berechnung der Gesamtsummen (letzte Spalte)
    for konto in first_map:
        konto[6] = sum(filter(None, konto[2:6]))  # Summe aller nicht-None Werte


if __name__ == "__main__":
    source_directory = os.path.abspath("../../source_tot/")
    process_files(source_directory)

    # Ausgabe zur Kontrolle
    for row in first_map:
        print(row)
