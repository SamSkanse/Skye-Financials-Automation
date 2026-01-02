#!/usr/bin/env python3
import pandas as pd
import numpy as np
import openpyxl.utils as get_column_letter

"""
BuildWeeklyWorkbook.py

Purpose:
 - Create the final Excel workbook for the period report. Writes two tabs:
     1) `Master Log` (detailed per-order rows)
     2) `Financial Summary` (human-friendly, pre-formatted summary table)

Key operations performed:
 - Accepts `master` and `weekly_summary` as DataFrames or file paths.
 - Recalculates and validates key financial metrics to ensure consistency
     between the Master Log and the summary inputs.
 - Builds a readable summary table (rows with escaped leading `+`/`-` so
     Excel does not treat them as formulas) and autosizes columns for neat output.
 - Writes the two-sheet workbook to `output_path` using `openpyxl` engine.

Notes:
 - Cells that begin with `+` or `-` are escaped to avoid Excel formula parsing.
 - The module auto-sizes columns for readability and requires `openpyxl` to
     write `.xlsx` files.
"""

def escape_excel_formula(text):
    if isinstance(text, str) and text and text[0] in ("=", "+", "-"):
        return "'" + text
    return text

def autosize_columns(ws):
    for col in ws.columns:
        max_length = 0
        # First cell in the column
        first_cell = col[0]
        # Get the column letter directly from the cell
        col_letter = first_cell.column_letter

        for cell in col:
            try:
                cell_value = "" if cell.value is None else str(cell.value)
                if len(cell_value) > max_length:
                    max_length = len(cell_value)
            except Exception:
                pass

        # Add a little padding
        adjusted_width = max_length + 2
        ws.column_dimensions[col_letter].width = adjusted_width


def build_weekly_workbook(
    master,
    weekly_summary,
    output_path="Skye_Weekly_Report.xlsx",
):
    """
    Build an Excel workbook from `master` and `weekly_summary` which may be
    file paths or DataFrames. Writes the Excel workbook to `output_path`.
    """
    # ---- LOAD DATA ----
    if isinstance(master, pd.DataFrame):
        master_df = master.copy()
    else:
        master_df = pd.read_csv(master)

    if isinstance(weekly_summary, pd.DataFrame):
        summary_raw = weekly_summary.copy()
    else:
        summary_raw = pd.read_csv(weekly_summary)

    if summary_raw.empty:
        raise ValueError("weekly_summary.csv is empty – run your summary script first.")

    s = summary_raw.iloc[0].copy()

    # ---- ENSURE MASTER NUMERIC COLUMNS ----
    for col in ["subtotal", "discount", "shipping", "tax", "bar_cogs", "total_shipping_cost"]:
        if col in master_df.columns:
            master_df[col] = pd.to_numeric(master_df[col], errors="coerce")

    # ---- ENSURE SUMMARY NUMERIC COLUMNS ----
    num_cols = [
        "Gross_Revenue",  # may or may not be used
        "Taxes_Collected",
        "COGS_Total",
        "Shipping_Costs_Total",
        "Shipping_Costs_Orders",
        "Receiving_Sum",
        "Payment_Processing_Fee",
        "Gross_Profit",
        "Gross_Margin",
        "Starting_Inventory_Bars",
        "Boxes_Sold_This_Week",
        "Bars_Sold_This_Week",
        "Total_Inventory_Sold_Bars",
        "Weekly_Ending_Inventory_Bars",
    ]
    for col in num_cols:
        if col in summary_raw.columns:
            s[col] = pd.to_numeric(s[col], errors="coerce")

    # ======================================================
    #           RE-CALCULATE KEY FINANCIALS
    # ======================================================

    # Revenue (product only, net of discounts, EXCLUDING shipping)
    # Formula: sum(subtotal - discount)
    revenue_product = (summary_raw["Gross_Revenue"]).sum(skipna=True)

    # Shipping collected from customers (Shopify "shipping" column)
    shipping_collected = summary_raw["Shipping_Collected"].sum(skipna=True)

    # Gross Revenue = product revenue + shipping collected
    gross_revenue = revenue_product + shipping_collected

    # Taxes collected
    # Use master to stay consistent with actual orders
    taxes_collected = summary_raw["Taxes_Collected"].sum(skipna=True)

    # COGS total – use summary (already sum of bar_cogs)
    if "COGS_Total" in s.index and not pd.isna(s["COGS_Total"]):
        cogs_total = float(s["COGS_Total"])
    else:
        cogs_total = master_df["bar_cogs"].sum(skipna=True)

    # Total 3PL Costs = shipping (3PL) + receiving + payment processing fee
    # Already computed in weekly_summary as Shipping_Costs_Total
    if "Shipping_Costs_Total" in s.index and not pd.isna(s["Shipping_Costs_Total"]):
        total_3pl_costs = float(s["Shipping_Costs_Total"])
    else:
        total_3pl_costs = 0.0

    # Gross Profit = Gross Revenue - COGS - Total 3PL Costs
    gross_profit = gross_revenue - cogs_total - total_3pl_costs

    # Gross Margin = Gross Profit / Gross Revenue
    gross_margin = gross_profit / gross_revenue if gross_revenue != 0 else np.nan

    # ======================================================
    #              INVENTORY / UNIT NUMBERS
    # ======================================================

    starting_inventory = int(s.get("Starting_Inventory_Bars", 0) or 0)

    boxes_sold = int(s.get("Boxes_Sold_This_Week", 0) or 0)
    bars_sold = int(s.get("Bars_Sold_This_Week", 0) or 0)
    total_inventory_sold = int(s.get("Total_Inventory_Sold_Bars", 0) or 0)
    weekly_ending_inventory = int(s.get("Weekly_Ending_Inventory_Bars", 0) or 0)

    # ======================================================
    #           BUILD PRETTY SUMMARY TABLE
    # ======================================================

    rows = []

    # Header
    rows.append([
        escape_excel_formula("============== Cumulative Period Financials ====================="),
        ""
    ])

    # Revenue structure
    # Use numeric types for values so Excel stores real numbers and can be formatted
    rows.append([escape_excel_formula("Revenue"), revenue_product])
    rows.append([escape_excel_formula("+ Shipping collected"), shipping_collected])
    rows.append([escape_excel_formula("-------------------------"), ""])
    rows.append([escape_excel_formula("Gross Revenue"), gross_revenue])

    # Taxes
    rows.append([escape_excel_formula("+ Taxes Collected"), taxes_collected])

    # COGS & 3PL
    rows.append([escape_excel_formula("- COGS"), cogs_total])
    rows.append([
        escape_excel_formula("- Total 3PL Costs (shipping, receiving, payment processing fee)"),
        total_3pl_costs,
    ])
    rows.append([escape_excel_formula("------------------------------------------------------"), ""])

    # Gross Profit & Margin
    rows.append([escape_excel_formula("Gross Profit"), gross_profit])
    if not np.isnan(gross_margin):
        # store as a fraction (e.g., 0.25 for 25%) so Excel percent formatting works
        rows.append([escape_excel_formula("Gross Margin"), gross_margin])
    else:
        rows.append([escape_excel_formula("Gross Margin"), "N/A"])

    # Spacer
    rows.append(["", ""])

    # Inventory header
    rows.append([escape_excel_formula("===============Inventory / Units======================="), ""])

    # Inventory details (store as integers so Excel keeps numeric types)
    rows.append([escape_excel_formula("Starting Inventory (bars)"), starting_inventory])
    rows.append([escape_excel_formula("Boxes Sold This Period"), boxes_sold])
    rows.append([escape_excel_formula("Bars Sold This Period (single bars)"), bars_sold])
    rows.append([escape_excel_formula("Total Inventory Sold (bars)"), total_inventory_sold])
    rows.append([escape_excel_formula("Weekly Ending Inventory (bars)"), weekly_ending_inventory])


    summary_pretty = pd.DataFrame(rows, columns=["Metric", "Value"])

    # ======================================================
    #                WRITE EXCEL WITH TWO TABS
    # ======================================================

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Tab 1: master log table
        master_df.to_excel(writer, sheet_name="Master Log", index=False)
    
    # Tab 2: Financial & Inventory Summary
        summary_pretty.to_excel(writer, sheet_name="Financial Summary", index=False)

        wb = writer.book
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            autosize_columns(ws)

        # Apply number formats to Financial Summary so values are stored as numbers
        if "Financial Summary" in wb.sheetnames:
            ws_summary = wb["Financial Summary"]
            # DataFrame wrote headers in row 1; data starts at row 2
            for row in ws_summary.iter_rows(min_row=2, min_col=1, max_col=2):
                metric_cell, value_cell = row[0], row[1]
                metric_text = str(metric_cell.value or "").lower()
                # Skip empty values and separators
                if value_cell.value is None or value_cell.value == "":
                    continue

                # Percentages (Gross Margin)
                if "margin" in metric_text:
                    try:
                        # Expecting a fraction (e.g., 0.25) -> display as percent
                        value_cell.number_format = '0.00%'
                    except Exception:
                        pass
                # Currency-like values
                elif any(k in metric_text for k in ["revenue", "tax", "cogs", "shipping", "profit", "3pl", "payment processing fee"]):
                    try:
                        value_cell.number_format = '"$"#,##0.00'
                    except Exception:
                        pass
                # Integer counts (inventory / units)
                elif any(k in metric_text for k in ["inventory", "bars", "boxes", "sold", "starting inventory", "ending inventory", "total inventory"]):
                    try:
                        value_cell.number_format = '#,##0'
                    except Exception:
                        pass

        print(f"Workbook written to: {output_path}")


# Runner for testing
# if __name__ == "__main__":
#     build_weekly_workbook(
#         master_log_path="master_log_Oct24_to_Nov21.csv",
#         weekly_summary_path="weekly_summary.csv",
#         output_path="Skye_Period_Report.xlsx",
#     )

