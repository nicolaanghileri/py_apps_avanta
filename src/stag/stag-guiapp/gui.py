# gui.py
import tkinter as tk
from tkinter import filedialog, StringVar
from main_algo import main_algo
from tkinter import ttk
import sv_ttk

source_directory = ""
output_directory = ""

def ordner_auswählen(label_ergebnis):
    global source_directory
    ordner_pfad = filedialog.askdirectory(title="Wähle Source Ordner aus")
    if ordner_pfad:
        label_ergebnis.config(text=f"Ausgewählter Ordner:\n{ordner_pfad}")
        source_directory = ordner_pfad
    else:
        label_ergebnis.config(text="Kein Ordner ausgewählt.")

def ordner_auswählen_output(label_ergebnis_output):
    global output_directory
    ordner_pfad = filedialog.askdirectory(title="Wähle Output Ordner aus")
    if ordner_pfad:
        label_ergebnis_output.config(text=f"Ausgewählter Ordner:\n{ordner_pfad}")
        output_directory = ordner_pfad
    else:
        label_ergebnis_output.config(text="Kein Ordner ausgewählt.")

def start_programm(selected_option, label_status):
    global source_directory, output_directory
    main_algo(source_directory, selected_option, label_status)

def run_gui():
    root = tk.Tk()
    sv_ttk.set_theme("dark")
    root.title("Avanta / Steigenberger Hotels AG")
    root.geometry("550x500")

    label_überschrift = ttk.Label(root, text="Monthly Report generator", font=("Helvetica", 16))
    label_überschrift.pack(pady=10)

    selected_option = StringVar()
    options = [
        "Monthly Report - P&L",
        "Monthly Report - Revenue",
        "Import Credit Card Payments",
        "Consolidation"
    ]
    default_value = "Monthly Report - P&L"
    dropdown = ttk.OptionMenu(root, selected_option, default_value, *options)
    dropdown.pack(pady=10)

    label_ergebnis = ttk.Label(root, text="", font=("Helvetica", 10), wraplength=350)
    label_ergebnis.pack(pady=10)
    button_auswahl = ttk.Button(root, text="Ordner auswählen (Source)", command=lambda: ordner_auswählen(label_ergebnis))
    button_auswahl.pack(pady=10)

    # Status-Label als tk.Label (damit fg unterstützt wird)
    label_status = tk.Label(root, text="", font=("Helvetica", 25))
    label_status.pack(pady=30)
    button_start = ttk.Button(root, text="START", command=lambda: start_programm(selected_option.get(), label_status))
    button_start.pack(pady=10)

    footer_frame = ttk.Frame(root)
    footer_frame.pack(side="bottom", fill="x")
    footer_label = ttk.Label(footer_frame, text="Nicola Anghileri - Avanta Group AG", font=("Helvetica", 9))
    footer_label.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    run_gui()
