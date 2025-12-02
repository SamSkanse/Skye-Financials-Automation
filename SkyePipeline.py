#!/usr/bin/env python3
import os
from MasterLogCreation import build_master_log
from WeeklySummaryCreator import build_weekly_summary, get_float_input, get_int_input
from BuildWeeklyWorkbook import build_weekly_workbook

def main():
    print("=== Skye Period Report Pipeline ===")

    # ---- Get file paths from user (with sensible defaults) ----
    orders_file = input("Path to Shopify orders CSV [orders_export_1 (2).csv]: ").strip()
    if not orders_file:
        orders_file = "orders_export_1 (2).csv"

    threepl_file = input("Path to 3PL Excel file [Skye Performance...xlsx]: ").strip()
    if not threepl_file:
        threepl_file = "Skye Performance 11.17.25 to 11.23.25.xlsx"

    master_log_output = input(
        "Output master log CSV [master_log_all_orders.csv]: "
    ).strip()
    if not master_log_output:
        master_log_output = "master_log_all_orders.csv"

    weekly_summary_output = "weekly_summary.csv"  # can parameterize later if you like

    report_output = input(
        "Output Excel report [Skye_Period_Report.xlsx]: "
    ).strip()
    if not report_output:
        report_output = "Skye_Period_Report.xlsx"

    # ---- Get financial inputs once ----
    print("\n=== Period Inputs ===")
    payment_processing_fee = get_float_input(
        "Enter payment processing fee for the period (e.g. 123.45): ", decimals=2
    )
    starting_inventory = get_int_input(
        "Enter starting inventory (in bars, whole number): "
    )

    # ---- STEP 1: build master log ----
    print("\n[1/3] Building master log...")
    build_master_log(orders_file, threepl_file, master_log_output)

    # ---- STEP 2: build weekly summary ----
    print("\n[2/3] Building weekly summary...")
    build_weekly_summary(
        master_log_output,
        threepl_file,
        weekly_summary_output,
        payment_processing_fee=payment_processing_fee,
        starting_inventory=starting_inventory,
    )

    # ---- STEP 3: build final Excel report ----
    print("\n[3/3] Building final Excel report...")
    build_weekly_workbook(
        master_log_path=master_log_output,
        weekly_summary_path=weekly_summary_output,
        output_path=report_output,
    )

    print(f"\n✅ Done! Final report written to: {os.path.abspath(report_output)}")


if __name__ == "__main__":
    main()
