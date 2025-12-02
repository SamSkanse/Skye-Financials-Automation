#!/usr/bin/env python3
import pandas as pd
import numpy as np

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
    master_path,
    threepl_path,
    output_path,
    payment_processing_fee=None,
    starting_inventory=None,
):
    # ---- LOAD DATA ----
    master = pd.read_csv(master_path)
    threepl = pd.read_excel(threepl_path)

    # Ensure key numeric columns are numeric
    for col in ["total", "tax", "bar_cogs", "total_shipping_cost",
                "line_item_quantity", "total_bars_sold"]:
        if col in master.columns:
            master[col] = pd.to_numeric(master[col], errors="coerce")

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
    # Gross Revenue (includes shipping) = Sum of (Total - Tax - shipping)
    gross_revenue = (master["total"] - master["tax"] - master["shipping"]).sum(skipna=True)

    shipping_collected = master["shipping"].sum(skipna=True)

    # Taxes Collected = Sum of taxes
    taxes_collected = master["tax"].sum(skipna=True)

    # ---- COGS ----
    cogs_total = master["bar_cogs"].sum(skipna=True)

    # ---- SHIPPING COSTS ----
    # 1) Sum of total_shipping_cost from master log (3PL per order)
    shipping_costs_orders = master["total_shipping_cost"].sum(skipna=True)

    # 2) Sum of Receiving column(s) from 3PL spreadsheet (for Shipment Orders)
    shipments = threepl.copy()

    # Try several likely column names for "Receiving"
    candidate_receiving_cols = [
        "Receiving"
    ]
    receiving_cols = [c for c in candidate_receiving_cols if c in shipments.columns]

    if receiving_cols:
        receiving_df = shipments[receiving_cols].apply(
            pd.to_numeric, errors="coerce"
        )
        receiving_sum = receiving_df.sum().sum()
    else:
        receiving_sum = 0.0

    # 3) Add payment processing fee (user input)
    shipping_costs_total = shipping_costs_orders + receiving_sum + payment_processing_fee

    # ---- GROSS PROFIT & MARGIN ----
    gross_profit = gross_revenue + shipping_collected - cogs_total - shipping_costs_total
    gross_margin = gross_profit / (gross_revenue + shipping_collected) if gross_revenue != 0 else np.nan

    # ---- INVENTORY / SALES ----
    # Boxes sold this week: sum line_item_quantity where box_or_bar == "box"
    boxes_sold = master.loc[
        master["box_or_bar"] == "box", "line_item_quantity"
    ].sum(skipna=True)

    # Bars sold this week: sum line_item_quantity where box_or_bar == "bar"
    bars_sold = master.loc[
        master["box_or_bar"] == "bar", "line_item_quantity"
    ].sum(skipna=True)

    # Total inventory sold (bars) = sum total_bars_sold from master log
    total_inventory_sold = master["total_bars_sold"].sum(skipna=True)

    # Weekly ending inventory (bars)
    weekly_ending_inventory = starting_inventory - total_inventory_sold

    # ---- BUILD SUMMARY ROW ----
    summary = {
        "Gross_Revenue": gross_revenue,
        "Shipping_Collected": shipping_collected,
        "Taxes_Collected": taxes_collected,
        "COGS_Total": cogs_total,
        "Shipping_Costs_Total": shipping_costs_total,
        "Shipping_Costs_Orders": shipping_costs_orders,
        "Receiving_Sum": receiving_sum,
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
    summary_df.to_csv(output_path, index=False)

    # ---- PRINT A NICE SUMMARY ----
    print("\n===== CUMULATIVE WEEKLY FINANCIALS (OVERALL) =====")
    print(f"Gross Revenue:           ${gross_revenue:,.2f}")
    print(f"Shipping Collected:           ${shipping_collected:,.2f}")
    print(f"Taxes Collected:         ${taxes_collected:,.2f}")
    print(f"COGS (total):            ${cogs_total:,.2f}")
    print(f"Shipping Costs (3PL): ${shipping_costs_orders:,.2f}")
    print(f"Receiving (3PL):         ${receiving_sum:,.2f}")
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

    print(f"\nWeekly summary written to: {output_path}")


if __name__ == "__main__":
    master_file = "master_log_Oct24_to_Nov21.csv"  # output from step 1
    threepl_file = "Skye Performance 11.17.25 to 11.23.25.xlsx"
    output_file = "weekly_summary.csv"

    build_weekly_summary(master_file, threepl_file, output_file)
