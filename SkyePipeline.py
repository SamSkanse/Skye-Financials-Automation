#!/usr/bin/env python3
import os
import re
from datetime import datetime
import pandas as pd
from skyepipeline_files.MasterLogCreation import build_master_log
from skyepipeline_files.WeeklySummaryCreator import build_weekly_summary, get_float_input, get_int_input
from skyepipeline_files.BuildWeeklyWorkbook import build_weekly_workbook

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

    # ---- input files (CHANGE EACH TIME) ----
    orders_file = "/Users/samskanse/desktop/calibrate_skye_12-15_12-25/orders_export_1-2.csv"
    if not orders_file:
        raise FileNotFoundError("Orders file path not provided.")
    if not os.path.isfile(orders_file):
        raise FileNotFoundError(f"Orders file not found: {orders_file}")

    threepl_file = "/Users/samskanse/desktop/calibrate_skye_12-15_12-25/Skye Performance 12.15.25 to 12.21.25.xlsx"
    if not threepl_file:
        raise FileNotFoundError("3PL file path not provided.")
    if not os.path.isfile(threepl_file):
        raise FileNotFoundError(f"3PL file not found: {threepl_file}")

    
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


    # --- final report output (CHANGE EACH TIME) ----
    if start_date is not None and end_date is not None:
        report_output = f"/Users/samskanse/desktop/Skye_Period_Report_{start_date.strftime('%Y-%m-%d')}_to_{end_date.strftime('%Y-%m-%d')}.xlsx"
    else:
        report_output = "/Users/samskanse/desktop/skye_period_reports/Skye_Period_Report.xlsx"

    # ---- financial inputs (Ask for) ----
    print("\n=== Period Inputs ===")
    payment_processing_fee = get_float_input(
        "Enter payment processing fee for the period (e.g. 123.45): ", decimals=2
    )
    starting_inventory = get_int_input(
        "Enter starting inventory (in bars, whole number): "
    )

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
    )

    print(f"\n✅ Done! Final report written to: {os.path.abspath(report_output)}")


if __name__ == "__main__":
    main()
