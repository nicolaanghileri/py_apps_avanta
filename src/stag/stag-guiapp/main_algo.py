# main_algo.py
from pathlib import Path
from datetime import date
from logic_pl import process_pl_report, process_revenue_report
from logic_cc import import_credit_card_payments
from logic_con import consolidation

def main_algo(source_dir, selected_option, label_status=None):
    # Falls "source_dir" noch ein String ist, mach ihn zu einem Path:
    source_dir = Path(source_dir)

    if not source_dir.is_dir():
        if label_status:
            label_status.config(text="Ungültiger Source Pfad", fg="red")
        return

    today = date.today().strftime("%d-%m-%Y")

    gen_output_dir = source_dir.parent / "output"
    gen_output_dir.mkdir(parents=True, exist_ok=True)

    if selected_option == "Monthly Report - P&L":
        final_output_dir = gen_output_dir / f"{today} P&L"
        final_output_dir.mkdir(parents=True, exist_ok=True)
        process_pl_report(source_dir, final_output_dir, label_status)

    elif selected_option == "Monthly Report - Revenue":
        final_output_dir = gen_output_dir / f"{today} Revenue"
        final_output_dir.mkdir(parents=True, exist_ok=True)
        process_revenue_report(source_dir, final_output_dir, label_status)

    elif selected_option == "Import Credit Card Payments":
        final_output_dir = gen_output_dir / f"{today} CC"
        final_output_dir.mkdir(parents=True, exist_ok=True)
        import_credit_card_payments(source_dir, final_output_dir, label_status)

    elif selected_option == "Consolidation":
        final_output_dir = gen_output_dir / f"{today} Conso"
        final_output_dir.mkdir(parents=True, exist_ok=True)
        consolidation(source_dir, final_output_dir, label_status)
        
    else:
        if label_status:
            label_status.config(text="Unbekannte Option!", fg="red")
