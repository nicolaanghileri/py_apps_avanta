# gui.py

import tkinter as tk
from tkinter import filedialog
from logic import main_algo

# Globale Variablen für die Pfade
source_directory = ""
output_directory = ""

def ordner_auswählen(label_ergebnis):
    """
    Öffnet einen Dialog, um den Source-Ordner auszuwählen.
    """
    global source_directory
    ordner_pfad = filedialog.askdirectory(title="Wähle Source Ordner aus")
    if ordner_pfad:
        label_ergebnis.config(text=f"Ausgewählter Ordner:\n{ordner_pfad}")
        source_directory = ordner_pfad
    else:
        label_ergebnis.config(text="Kein Ordner ausgewählt.")


def ordner_auswählen_output(label_ergebnis_output):
    """
    Öffnet einen Dialog, um den Output-Ordner auszuwählen.
    """
    global output_directory
    ordner_pfad = filedialog.askdirectory(title="Wähle Output Ordner aus")
    if ordner_pfad:
        label_ergebnis_output.config(text=f"Ausgewählter Ordner:\n{ordner_pfad}")
        output_directory = ordner_pfad
    else:
        label_ergebnis_output.config(text="Kein Ordner ausgewählt.")


def start_programm(label_status):
    """
    Ruft die main_algo aus logic.py auf und übergibt den
    aktuell ausgewählten source- und output-Pfad. label_status
    dient dazu, das Ergebnis in der GUI anzuzeigen.
    """
    main_algo(source_directory, output_directory, label_status)


def run_gui():
    """
    Hier wird das Hauptfenster (Tk) erstellt und alle GUI-Elemente platziert.
    Am Ende starten wir die mainloop.
    """
    root = tk.Tk()
    root.title("Avanta / Steigenberger Hotels AG")
    root.geometry("550x500")

    # Überschrift
    label_überschrift = tk.Label(root, text="Monthly Report generator", font=("Helvetica", 16))
    label_überschrift.pack(pady=10)

    # Label für Quelle-Ordner
    label_ergebnis = tk.Label(root, text="", font=("Helvetica", 10), wraplength=350)
    label_ergebnis.pack(pady=10)

    # Button zur Ordnerauswahl (Source)
    button_auswahl = tk.Button(
        root, 
        text="Ordner auswählen (Source)", 
        command=lambda: ordner_auswählen(label_ergebnis)
    )
    button_auswahl.pack(pady=10)

    # Label für Output-Ordner
    label_ergebnis_output = tk.Label(root, text="", font=("Helvetica", 10), wraplength=350)
    label_ergebnis_output.pack(pady=10)

    # Button zur Ordnerauswahl (Output)
    button_auswahl_output = tk.Button(
        root, 
        text="Ordner auswählen (Output)", 
        command=lambda: ordner_auswählen_output(label_ergebnis_output)
    )
    button_auswahl_output.pack(pady=10)

    # Button zum Start des Programms
    label_status = tk.Label(root, text="", font=("Helvetica", 25))
    label_status.pack(pady=30)

    button_start = tk.Button(root, text="START", command=lambda: start_programm(label_status))
    button_start.pack(pady=10)

    # Footer
    footer_frame = tk.Frame(root, bg="lightgray")
    footer_frame.pack(side="bottom", fill="x")

    footer_label = tk.Label(
        footer_frame,
        text="nicola.anghileri@avanta.ch - Avanta Group AG",
        bg="lightgray",
        font=("Helvetica", 9)
    )
    footer_label.pack(pady=5)

    # GUI starten
    root.mainloop()
