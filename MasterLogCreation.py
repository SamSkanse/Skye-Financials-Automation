#!/usr/bin/env python3
import pandas as pd
import numpy as np
from pathlib import Path

# ---- CONSTANTS ----
# Per-bar COGS from: 39,891.91 / 15,848
PER_BAR_COGS = 39891.91 / 15848  # ≈ 2.517...


# ---- HELPER FUNCTIONS ----

def classify_shopify_item(price):
    """
    For Shopify rows:
    box = line item price > 20
    bar  = line item price < 5
    otherwise None
    """
    try:
        p = float(price)
    except (TypeError, ValueError):
        return None
    if p > 20:
        return "box"
    elif p < 5:
        return "bar"
    else:
        return None


def classify_sample_item(total_price, qty):
    """
    For free sample 3PL rows:
    box = total price / quantity > 20
    bar = total price / quantity < 10
    otherwise None
    """
    try:
        total_price = float(total_price)
        qty = float(qty)
        if qty <= 0:
            return None
        unit = total_price / qty
    except (TypeError, ValueError, ZeroDivisionError):
        return None

    if unit > 20:
        return "box"
    elif unit < 10:
        return "bar"
    else:
        return None


def compute_bars_sold(item_type, qty):
    """
    7 bars per box.
    """
    if pd.isna(qty):
        return 0
    if item_type == "box":
        return qty * 7
    elif item_type == "bar":
        return qty
    else:
        return 0


def compute_bar_cogs(item_type, qty):
    bars = compute_bars_sold(item_type, qty)
    return bars * PER_BAR_COGS


def compute_total_shipping(row):
    """
    Total shipping cost (3PL):
    Handling Fee + Total Shipping Cost + Packaging

    If ALL three are missing, return NaN (blank in CSV).
    """
    vals = []
    for col in ["Handling Fee", "Total Shipping Cost", "Packaging"]:
        v = pd.to_numeric(row.get(col, np.nan), errors="coerce")
        vals.append(v)

    # If everything is NaN (no 3PL match), leave blank
    if all(pd.isna(v) for v in vals):
        return np.nan

    # Otherwise, sum the non-NaN pieces
    return sum(v for v in vals if pd.notna(v))


# ---- CORE LOGIC ----

def build_master_log(orders_path, threepl_path, output_path):
    # Read inputs
    orders = pd.read_csv(orders_path)
    threepl = pd.read_excel(threepl_path)

    # Only care about shipment rows in 3PL sheet
    shipments = threepl[threepl["Type"] == "Shipment Order"].copy()

    # ---- ALL SHOPIFY ORDERS ----
    shopify = orders.copy()

    # 3PL rows that have a Store Order Number (normal orders)
    threepl_orders = shipments[shipments["Store Order Number"].notna()].copy()

    # Merge Shopify line items with matching 3PL shipment row (if any)
    merged = shopify.merge(
        threepl_orders[
            ["Store Order Number", "Handling Fee", "Total Shipping Cost", "Packaging"]
        ],
        left_on="Name",
        right_on="Store Order Number",
        how="left",     # <-- key change: keep ALL Shopify orders
    )

    # Clean / compute fields for normal orders
    merged["line_item_quantity"] = pd.to_numeric(
        merged["Lineitem quantity"], errors="coerce"
    )
    merged["line_item_price"] = pd.to_numeric(
        merged["Lineitem price"], errors="coerce"
    )

    merged["box_or_bar"] = merged["line_item_price"].apply(classify_shopify_item)
    merged["total_bars_sold"] = merged.apply(
        lambda r: compute_bars_sold(r["box_or_bar"], r["line_item_quantity"]), axis=1
    )
    merged["bar_cogs"] = merged.apply(
        lambda r: compute_bar_cogs(r["box_or_bar"], r["line_item_quantity"]), axis=1
    )
    merged["total_shipping_cost"] = merged.apply(compute_total_shipping, axis=1)

    merged_rows = pd.DataFrame(
        {
            # Order Details
            "order_ID": merged["Name"],  # Shopify "Name"
            "order_date": merged["Paid at"],
            "email": merged["Email"],
            "box_or_bar": merged["box_or_bar"],
            "source": merged["Source"],
            "line_item_quantity": merged["line_item_quantity"],
            "total_bars_sold": merged["total_bars_sold"],
            "line_item_price": merged["line_item_price"],
            # Order Financials
            "subtotal": pd.to_numeric(merged["Subtotal"], errors="coerce"),
            "discount": pd.to_numeric(merged["Discount Amount"], errors="coerce"),
            "shipping": pd.to_numeric(merged["Shipping"], errors="coerce"),
            "tax": pd.to_numeric(merged["Taxes"], errors="coerce"),
            "total": pd.to_numeric(merged["Total"], errors="coerce"),
            # Costs
            "bar_cogs": merged["bar_cogs"],
            "total_shipping_cost": merged["total_shipping_cost"],
        }
    )

    # ---- FREE SAMPLE SHIPMENTS: 3PL rows with no Store Order Number ----

    samples = shipments[shipments["Store Order Number"].isna()].copy()

    samples["line_item_quantity"] = pd.to_numeric(
        samples["Total Quantity"], errors="coerce"
    )
    samples["unit_price"] = (
        pd.to_numeric(samples["Total Price"], errors="coerce")
        / samples["line_item_quantity"]
    )
    samples["box_or_bar"] = samples.apply(
        lambda r: classify_sample_item(r["Total Price"], r["Total Quantity"]), axis=1
    )
    samples["total_bars_sold"] = samples.apply(
        lambda r: compute_bars_sold(r["box_or_bar"], r["line_item_quantity"]), axis=1
    )
    samples["bar_cogs"] = samples.apply(
        lambda r: compute_bar_cogs(r["box_or_bar"], r["line_item_quantity"]), axis=1
    )
    samples["total_shipping_cost"] = samples.apply(compute_total_shipping, axis=1)

    sample_rows = pd.DataFrame(
        {
            # For free samples, fill only what we can from 3PL
            "order_ID": samples["Order Code"],  # no Shopify order number
            "order_date": samples["Actual Shipment Date"],
            "email": "FREE SAMPLE BOX",
            "box_or_bar": samples["box_or_bar"],
            "source": "free_sample",
            "line_item_quantity": samples["line_item_quantity"],
            "total_bars_sold": samples["total_bars_sold"],
            "line_item_price": samples["unit_price"],
            "subtotal": np.nan,  # not provided in spec for free samples
            "discount": pd.to_numeric(samples["Custom Discount"], errors="coerce"),
            "shipping": np.nan,  # Shopify shipping blank; shipping captured in total_shipping_cost
            "tax": pd.to_numeric(samples["Total Tax"], errors="coerce"),
            "total": np.nan,
            "bar_cogs": samples["bar_cogs"],
            "total_shipping_cost": samples["total_shipping_cost"],
        }
    )

    # ---- FINAL MASTER LOG ----

    master = pd.concat([merged_rows, sample_rows], ignore_index=True)
    master.to_csv(output_path, index=False)
    print(f"Master log written to: {output_path}")


if __name__ == "__main__":
    # Update these paths as needed
    orders_file = "orders_export_1.csv"
    threepl_file = "Skye Performance 11.17.25 to 11.23.25.xlsx"
    output_file = "master_log_all_orders.csv"

    build_master_log(orders_file, threepl_file, output_file)
