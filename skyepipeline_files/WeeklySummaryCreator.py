#!/usr/bin/env python3
import pandas as pd
import numpy as np
from pathlib import Path

"""
WeeklySummaryCreator.py

Purpose:
 - Compute the weekly Financial & Inventory summary from the Master Log
     and 3PL data. Produces a one-row summary DataFrame with the key metrics
     required for the period report.

Key operations performed:
 - Accepts `master` and `threepl` as DataFrames or file paths.
 - Normalizes numeric columns and prompts for runtime inputs when not supplied:
     payment processing fee and starting inventory (bars).
 - Calculates gross revenue, shipping collected, taxes collected, COGS,
     3PL shipping/receiving costs, total shipping costs (including payment fee),
     gross profit, and gross margin.
 - Computes inventory movement (boxes/bars sold) and excludes rows labeled
     `source == 'sales_team'` from boxes/bars counts so internal/GTM sendouts do
     not skew sales metrics.
 - Returns a one-row `summary_df` and optionally writes to CSV/XLSX.

Notes:
 - The function is designed to be used programmatically (returns DataFrame)
     and interactively (prompts for missing inputs).
 - Excel output requires `openpyxl` when writing `.xlsx` files.
"""

def get_float_input(prompt, decimals=2):
    while True:
        s = input(prompt).strip()
        try:
            value = float(s)
            return round(value, decimals)
        except ValueError:
            print("Please enter a valid number (e.g. 123.45).")

def get_int_input(prompt):
    while True:
        s = input(prompt).strip()
        try:
            value = int(float(s))
            return value
        except ValueError:
            print("Please enter a whole number (e.g. 10000).")

def build_weekly_summary(
    master,
    threepl,
    output_path=None,
    payment_processing_fee=None,
    starting_inventory=None,
):
    """
    Build weekly summary. `master` and `threepl` may be file paths or DataFrames.

    Returns a pandas.DataFrame (summary_df). If `output_path` is provided the
    CSV/Excel will also be written.
    """
    # ---- LOAD DATA ----
    if isinstance(master, pd.DataFrame):
        master_df = master.copy()
    else:
        master_df = pd.read_csv(master)

    if isinstance(threepl, pd.DataFrame):
        threepl_df = threepl.copy()
    else:
        threepl_df = pd.read_excel(threepl)

    # Ensure key numeric columns are numeric
    for col in ["total", "tax", "bar_cogs", "total_shipping_cost",
                "line_item_quantity", "total_bars_sold"]:
        if col in master_df.columns:
            master_df[col] = pd.to_numeric(master_df[col], errors="coerce")

    # ---- ASK USER INPUTS ----
    if payment_processing_fee is None:
        payment_processing_fee = get_float_input(
            "Enter payment processing fee for the week (e.g. 123.45): ",
            decimals=2
        )

    if starting_inventory is None:
        starting_inventory = get_int_input(
            "Enter starting inventory (in bars, whole number): "
        )

    # ---- GROSS REVENUE & TAXES ----
    gross_revenue = (master_df["total"] - master_df["tax"] - master_df["shipping"]).sum(skipna=True)
    shipping_collected = master_df["shipping"].sum(skipna=True)
    taxes_collected = master_df["tax"].sum(skipna=True)

    # ---- COGS ----
    cogs_total = master_df["bar_cogs"].sum(skipna=True)

    # ---- SHIPPING COSTS ----
    # Sum per-order shipping from the master frame. Some sample rows used
    shipping_costs_orders = 0.0
    for col in ["total_shipping_cost"]:
        if col in master_df.columns:
            shipping_costs_orders += master_df[col].sum(skipna=True)

    shipments = threepl_df.copy()

    # ---- Extra 3PL rows (e.g., Handling, Receiving, Freight, Storage) ----
    # Sum any cost-like columns for rows where Type != 'Shipment Order'.
    # These columns can include: Handling Fee, Total Shipping Cost, LTL Freight,
    # Packaging, Label Fee, Receiving, Returns, Storage. If present, they will
    # be summed per-row and added to the period 3PL extra costs.
    extra_cols_candidates = [
        "Handling Fee",
        "Total Shipping Cost",
        "LTL Freight",
        "Packaging",
        "Label Fee",
        "Receiving",
        "Returns",
        "Storage",
    ]
    ship_cols = [c for c in extra_cols_candidates if c in shipments.columns]
    extra_shipping_sum = 0.0
    if "Type" in shipments.columns and ship_cols:
        other_mask = ~shipments["Type"].astype(str).str.lower().eq("shipment order")
        if other_mask.any():
            extra_df = shipments.loc[other_mask, ship_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
            # Sum across the row then across rows
            extra_shipping_sum = extra_df.sum(axis=1).sum()

    shipping_costs_total = shipping_costs_orders + payment_processing_fee + extra_shipping_sum

    # ---- GROSS PROFIT & MARGIN ----
    gross_profit = gross_revenue + shipping_collected - cogs_total - shipping_costs_total
    gross_margin = gross_profit / (gross_revenue + shipping_collected) if gross_revenue != 0 else np.nan

    # ---- INVENTORY / SALES ----
    # Exclude sales-team samples from boxes/bars sold counts
    if "source" in master_df.columns:
        boxes_sold = master_df.loc[
            (master_df["box_or_bar"] == "box") & (master_df["source"] != "sales_team"),
            "line_item_quantity",
        ].sum(skipna=True)

        bars_sold = master_df.loc[
            (master_df["box_or_bar"] == "bar") & (master_df["source"] != "sales_team"),
            "line_item_quantity",
        ].sum(skipna=True)
    else:
        boxes_sold = master_df.loc[master_df["box_or_bar"] == "box", "line_item_quantity"].sum(skipna=True)
        bars_sold = master_df.loc[master_df["box_or_bar"] == "bar", "line_item_quantity"].sum(skipna=True)
    # Exclude GTM/sales sendouts from total inventory sold
    if "exclude_from_bars_sold" in master_df.columns:
        total_inventory_sold = master_df.loc[master_df["exclude_from_bars_sold"] == False, "total_bars_sold"].sum(skipna=True)
    else:
        total_inventory_sold = master_df["total_bars_sold"].sum(skipna=True)
    weekly_ending_inventory = starting_inventory - total_inventory_sold

    # ---- BUILD SUMMARY ROW ----
    summary = {
        "Gross_Revenue": gross_revenue,
        "Shipping_Collected": shipping_collected,
        "Taxes_Collected": taxes_collected,
        "COGS_Total": cogs_total,
        "Shipping_Costs_Total": shipping_costs_total,
        "Shipping_Costs_Orders": shipping_costs_orders,
        "3PL_Extra_Costs": extra_shipping_sum,
        "Payment_Processing_Fee": payment_processing_fee,
        "Gross_Profit": gross_profit,
        "Gross_Margin": gross_margin,
        "Starting_Inventory_Bars": starting_inventory,
        "Boxes_Sold_This_Week": boxes_sold,
        "Bars_Sold_This_Week": bars_sold,
        "Total_Inventory_Sold_Bars": total_inventory_sold,
        "Weekly_Ending_Inventory_Bars": weekly_ending_inventory,
    }

    summary_df = pd.DataFrame([summary])
    if output_path:
        out_path = Path(output_path)
        suffix = out_path.suffix.lower()
        if suffix in (".xlsx", ".xls"):
            try:
                summary_df.to_excel(output_path, index=False)
            except ImportError:
                raise ImportError("Writing Excel files requires 'openpyxl' (pip install openpyxl).")
        else:
            summary_df.to_csv(output_path, index=False)

    # ---- PRINT A NICE SUMMARY ----
    print("\n===== CUMULATIVE WEEKLY FINANCIALS (OVERALL) =====")
    print(f"Gross Revenue:           ${gross_revenue:,.2f}")
    print(f"Shipping Collected:           ${shipping_collected:,.2f}")
    print(f"Taxes Collected:         ${taxes_collected:,.2f}")
    print(f"COGS (total):            ${cogs_total:,.2f}")
    print(f"Shipping Costs (3PL): ${shipping_costs_orders:,.2f}")
    print(f"3PL Extra Costs:         ${extra_shipping_sum:,.2f}")
    print(f"Payment Processing Fee:  ${payment_processing_fee:,.2f}")
    print(f"Total Shipping Costs:    ${shipping_costs_total:,.2f}")
    print(f"Gross Profit:            ${gross_profit:,.2f}")
    if not np.isnan(gross_margin):
        print(f"Gross Margin:            {gross_margin*100:,.2f}%")
    else:
        print("Gross Margin:            N/A (Gross Revenue is 0)")

    print("\n===== INVENTORY / UNITS =====")
    print(f"Starting Inventory (bars):        {starting_inventory:,}")
    print(f"Boxes Sold This Week:             {int(boxes_sold)}")
    print(f"Bars Sold This Week (single bars):{int(bars_sold)}")
    print(f"Total Inventory Sold (bars):      {int(total_inventory_sold)}")
    print(f"Weekly Ending Inventory (bars):   {int(weekly_ending_inventory)}")

    if output_path:
        print(f"\nWeekly summary written to: {output_path}")

    return summary_df

# Runner for testing
# if __name__ == "__main__":
#     master_file = "master_log_Oct24_to_Nov21.csv"  # output from step 1
#     threepl_file = "Skye Performance 11.17.25 to 11.23.25.xlsx"
#     output_file = "weekly_summary.csv"

#     build_weekly_summary(master_file, threepl_file, output_file)
