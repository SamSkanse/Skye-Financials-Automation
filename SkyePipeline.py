#!/usr/bin/env python3

import os
import re
from datetime import datetime
import pandas as pd
import sys
from skyepipeline_files.MasterLogCreation import build_master_log
from skyepipeline_files.WeeklySummaryCreator import build_weekly_summary, get_float_input, get_int_input
from skyepipeline_files.BuildWeeklyWorkbook import build_weekly_workbook

# --- File picker helpers ---
def pick_file(title="Select file", filetypes=(('All files', '*.*'),)):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return file_path

def pick_directory(title="Select output folder"):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    dir_path = filedialog.askdirectory(title=title)
    root.destroy()
    return dir_path


def get_period_inputs_ui(defaults=None):
    import tkinter as tk

    defaults = defaults or {}
    result = {}
    cancelled = {"flag": False}

    root = tk.Tk()
    root.title("Period Inputs")

    tk.Label(root, text="Starting inventory (bars):").grid(row=0, column=0, sticky="e", padx=6, pady=6)
    start_var = tk.StringVar(value=str(defaults.get("starting_inventory", "")))
    tk.Entry(root, textvariable=start_var).grid(row=0, column=1, padx=6, pady=6)

    tk.Label(root, text="Payment processing fee ($):").grid(row=1, column=0, sticky="e", padx=6, pady=6)
    fee_var = tk.StringVar(value=str(defaults.get("payment_processing_fee", "")))
    tk.Entry(root, textvariable=fee_var).grid(row=1, column=1, padx=6, pady=6)
    
    tk.Label(root, text="Total POS Bars given to sales members:").grid(row=2, column=0, sticky="e", padx=6, pady=6)
    tot_pos_var = tk.StringVar(value=str(defaults.get("tot_pos_bars", "")))
    tk.Entry(root, textvariable=tot_pos_var).grid(row=2, column=1, padx=6, pady=6)

    tk.Label(root, text="Bars to be sold (POS):").grid(row=3, column=0, sticky="e", padx=6, pady=6)
    pos_var = tk.StringVar(value=str(defaults.get("pos_bars", "")))
    tk.Entry(root, textvariable=pos_var).grid(row=3, column=1, padx=6, pady=6)

    def on_ok():
        result["starting_inventory"] = start_var.get().strip()
        result["payment_processing_fee"] = fee_var.get().strip()
        result["tot_pos_bars"] = tot_pos_var.get().strip()
        result["pos_bars"] = pos_var.get().strip()
        root.destroy()

    def on_cancel():
        # mark cancellation so caller can exit cleanly
        result.clear()
        cancelled["flag"] = True
        root.destroy()

    tk.Button(root, text="OK", command=on_ok).grid(row=4, column=0, pady=8)
    tk.Button(root, text="Cancel", command=on_cancel).grid(row=4, column=1, pady=8)

    root.resizable(False, False)
    root.mainloop()

    # If the user cancelled/closed the window, return None so caller can exit
    if cancelled["flag"]:
        return None

    # Parse and coerce values with safe fallbacks
    try:
        si = int(float(result.get("starting_inventory", 0))) if result.get("starting_inventory", "") != "" else None
    except Exception:
        si = None
    try:
        ppf = float(result.get("payment_processing_fee", 0)) if result.get("payment_processing_fee", "") != "" else None
    except Exception:
        ppf = None
    try:        
        tot_pos_bars = int(float(result.get("tot_pos_bars", 0))) if result.get("tot_pos_bars", "") != "" else None
    except Exception:
        tot_pos_bars = None
    try:
        posv = int(float(result.get("pos_bars", 0))) if result.get("pos_bars", "") != "" else None
    except Exception:
        posv = None

    return {"starting_inventory": si, "payment_processing_fee": ppf, "tot_pos_bars": tot_pos_bars, "pos_bars": posv}

"""
SkyePipeline.py

Purpose:
 - Top-level orchestration script for the Skye period report pipeline.

Key operations performed:
 - Validate input file paths for the Shopify orders CSV and 3PL (Calibrate) Excel file.
 - Attempt to infer the report date range from the 3PL filename or fall back
     to scanning date-like columns in the 3PL sheet.
 - Prompt the user for runtime inputs (payment processing fee and starting
     inventory) when running interactively.
 - Execute the three main pipeline steps in-memory:
         1) `build_master_log` from `skyepipeline_files.MasterLogCreation` (returns DataFrame)
         2) `build_weekly_summary` from `skyepipeline_files.WeeklySummaryCreator` (returns DataFrame)
         3) `build_weekly_workbook` from `skyepipeline_files.BuildWeeklyWorkbook` (writes the Excel workbook)
 - Write the final two-tab Excel report (Master Log + Financial Summary) to the
     inferred `report_output` path.

Notes:
 - This script is intended as a runnable convenience wrapper. The core logic
     is implemented in the modules under `skyepipeline_files/` so they can be
     imported and tested independently.
"""

def main():
    print("=== Skye Period Report Pipeline ===")


    # ---- input files (auto-open file pickers) ----
    orders_file = pick_file(title="Select Shopify Orders CSV", filetypes=[('CSV files', '*.csv'), ('All files', '*.*')])
    if not orders_file:
        print("No orders file selected. Exiting.")
        sys.exit(0)
    if not os.path.isfile(orders_file):
        print(f"Orders file not found: {orders_file}")
        sys.exit(1)

    threepl_file = pick_file(title="Select 3PL Excel file", filetypes=[('Excel files', ('*.xlsx', '*.xls')), ('All files', '*.*')])
    if not threepl_file:
        print("No 3PL file selected. Exiting.")
        sys.exit(0)
    if not os.path.isfile(threepl_file):
        print(f"3PL file not found: {threepl_file}")
        sys.exit(1)

    
    # Try to infer start/end from filename like: '... 11.17.25 to 11.23.25.xlsx'
    def _infer_date_range_from_filename(fname: str):
        m = re.search(r"(\d{1,2}\.\d{1,2}\.\d{2})\s*to\s*(\d{1,2}\.\d{1,2}\.\d{2})", fname, re.IGNORECASE)
        if not m:
            return None, None

        def _parse_mmddyy(s: str):
            month, day, yy = s.split('.')
            year = 2000 + int(yy)
            return datetime(year=int(year), month=int(month), day=int(day))

        try:
            return _parse_mmddyy(m.group(1)), _parse_mmddyy(m.group(2))
        except Exception:
            return None, None

    start_date, end_date = _infer_date_range_from_filename(threepl_file)

    # Fallback: read the 3PL excel and search for any column containing 'date'
    if start_date is None:
        try:
            df_3pl = pd.read_excel(threepl_file)
            date_cols = [c for c in df_3pl.columns if 'date' in str(c).lower()]
            if date_cols:
                dates = []
                for c in date_cols:
                    parsed = pd.to_datetime(df_3pl[c], errors='coerce')
                    dates.extend(parsed.dropna().tolist())
                if dates:
                    start_date = min(dates)
                    end_date = max(dates)
        except Exception:
            start_date = None



    # --- final report output (auto-open directory picker) ----
    output_dir = pick_directory(title="Select output folder for report")
    if not output_dir:
        print("No output folder selected. Exiting.")
        sys.exit(0)
    if not os.path.isdir(output_dir):
        print(f"Output folder not found: {output_dir}")
        sys.exit(1)
    # (No Finder reveal calls — selection will not be opened automatically)

    if start_date is not None and end_date is not None:
        report_output = os.path.join(output_dir, f"Skye_Period_Report_{start_date.strftime('%Y-%m-%d')}_to_{end_date.strftime('%Y-%m-%d')}.xlsx")
    else:
        report_output = os.path.join(output_dir, "Skye_Period_Report.xlsx")

    # ---- financial inputs (small UI) ----
    ui_vals = get_period_inputs_ui()
    if ui_vals is None:
        print("Period inputs dialog was cancelled. Exiting.")
        sys.exit(0)
    if ui_vals.get("payment_processing_fee") is None:
        payment_processing_fee = get_float_input(
            "Enter payment processing fee for the period (e.g. 123.45): ", decimals=2
        )
    else:
        payment_processing_fee = ui_vals.get("payment_processing_fee")

    if ui_vals.get("starting_inventory") is None:
        starting_inventory = get_int_input(
            "Enter starting inventory (in bars, whole number): "
        )
    else:
        starting_inventory = ui_vals.get("starting_inventory")

    # POS bars value (may be used by BuildWeeklyWorkbook)
    pos_bars_val_from_ui = ui_vals.get("pos_bars")
    total_sales_team_pos_bars_from_ui = ui_vals.get("tot_pos_bars")

    # ---- STEP 1: build master log (returns DataFrame) ----
    print("\n[1/3] Building master log (in-memory)...")
    master_df = build_master_log(orders_file, threepl_file, output_path=None)

    # ---- STEP 2: build weekly summary (returns DataFrame) ----
    print("\n[2/3] Building weekly summary (in-memory)...")
    summary_df = build_weekly_summary(
        master_df,
        threepl_file,
        output_path=None,
        payment_processing_fee=payment_processing_fee,
        starting_inventory=starting_inventory,
    )

    # ---- STEP 3: build final Excel report (writes workbook) ----
    print("\n[3/3] Building final Excel report...")
    build_weekly_workbook(
        master_df,
        summary_df,
        output_path=report_output,
        pos_bars=pos_bars_val_from_ui,
        tot_pos_bars=total_sales_team_pos_bars_from_ui,
    )

    print(f"\n✅ Done! Final report written to: {os.path.abspath(report_output)}")


if __name__ == "__main__":
    main()
