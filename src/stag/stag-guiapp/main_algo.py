# main_algo.py
import os
from logic_pl import process_pl_revenue
from logic_cc import import_credit_card_payments
from logic_con import consolidation

def main_algo(source_dir, output_dir, selected_option, label_status=None):
    if not os.path.isdir(source_dir) or not os.path.isdir(output_dir):
        if label_status:
            label_status.config(text="Ungültige Pfade", fg="red")
        return

    if selected_option in ["Monthly Report - P&L", "Monthly Report - Revenue"]:
        process_pl_revenue(source_dir, output_dir, label_status)
    elif selected_option == "Import Credit Card Payments":
        import_credit_card_payments(source_dir, output_dir, label_status)
    elif selected_option == "Consolidation":
        consolidation(source_dir, output_dir, label_status)
    else:
        if label_status:
            label_status.config(text="Unbekannte Option!", fg="red")
