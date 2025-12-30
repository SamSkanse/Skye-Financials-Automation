#!/usr/bin/env python3
import pandas as pd
import numpy as np
import re
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
    elif p < 6.5:
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

# ---- Outputs only Order log without 3pl (used for testing) ----

def orders_log_from_csv(orders_file, output_path=None):
    """
    Build an orders-only log from a Shopify orders CSV.

    Parameters
    - orders_file: path to the Shopify orders CSV
    - output_path: optional path to write the CSV output

    Returns a `pandas.DataFrame` with the same columns as the Shopify portion
    of `build_master_log` (order_ID, order_date, email, box_or_bar, source,
    line_item_quantity, total_bars_sold, line_item_price, subtotal, discount,
    shipping, tax, total, bar_cogs, total_shipping_cost).
    """
    orders = pd.read_csv(orders_file)

    shopify = orders.copy()

    shopify["line_item_quantity"] = pd.to_numeric(
        shopify.get("Lineitem quantity"), errors="coerce"
    )
    shopify["line_item_price"] = pd.to_numeric(
        shopify.get("Lineitem price"), errors="coerce"
    )

    shopify["box_or_bar"] = shopify["line_item_price"].apply(classify_shopify_item)
    shopify["total_bars_sold"] = shopify.apply(
        lambda r: compute_bars_sold(r["box_or_bar"], r["line_item_quantity"]),
        axis=1,
    )
    shopify["bar_cogs"] = shopify.apply(
        lambda r: compute_bar_cogs(r["box_or_bar"], r["line_item_quantity"]),
        axis=1,
    )

    # No 3PL merge here, so total_shipping_cost is blank (NaN)
    shopify_rows = pd.DataFrame(
        {
            "order_ID": shopify.get("Name"),
            "order_date": shopify.get("Paid at"),
            "email": shopify.get("Email"),
            "box_or_bar": shopify["box_or_bar"],
            "source": shopify.get("Source"),
            "line_item_quantity": shopify["line_item_quantity"],
            "total_bars_sold": shopify["total_bars_sold"],
            "line_item_price": shopify["line_item_price"],
            # Order Financials
            "subtotal": pd.to_numeric(shopify.get("Subtotal"), errors="coerce"),
            "discount": pd.to_numeric(shopify.get("Discount Amount"), errors="coerce"),
            "shipping": pd.to_numeric(shopify.get("Shipping"), errors="coerce"),
            "tax": pd.to_numeric(shopify.get("Taxes"), errors="coerce"),
            "total": pd.to_numeric(shopify.get("Total"), errors="coerce"),
            # Costs
            "bar_cogs": shopify["bar_cogs"],
            "total_shipping_cost": np.nan,
        }
    )

    if output_path:
        out_path = Path(output_path)
        suffix = out_path.suffix.lower()

        # If user asked for an Excel file, write with to_excel (requires openpyxl)
        if suffix in (".xlsx", ".xls"):
            try:
                shopify_rows.to_excel(output_path, index=False)
                print(f"Orders-only Excel log written to: {output_path}")
            except ImportError:
                raise ImportError(
                    "Writing Excel files requires 'openpyxl' (pip install openpyxl)."
                )
        else:
            # Default to CSV for .csv or unknown extensions
            shopify_rows.to_csv(output_path, index=False)
            print(f"Orders-only CSV log written to: {output_path}")

    return shopify_rows


# ---- CORE LOGIC ----

def build_master_log(orders_path, threepl_path, output_path=None):
    """
    Build the full master log from orders CSV and 3PL Excel.

    Returns a pandas.DataFrame. If `output_path` is provided the CSV
    will also be written to that path (backwards-compatible).
    """
    # Read inputs (accept either DataFrame or path)
    if isinstance(orders_path, pd.DataFrame):
        orders = orders_path.copy()
    else:
        orders = pd.read_csv(orders_path)

    if isinstance(threepl_path, pd.DataFrame):
        threepl = threepl_path.copy()
    else:
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

    # --- Detect sales-team / GTM shipments in 3PL samples ---
    # If the 3PL row has no Store Order Number and the Description mentions
    # GTM or Sales team keywords, mark it as a sales_team item so it won't
    # count toward sold inventory or financials.
    desc_col = next((c for c in samples.columns if "description" in c.lower()), None)
    if desc_col:
        desc_series = samples[desc_col].fillna("").astype(str).str.lower()
        keywords = ["gtm", "sales team", "sales_team", "gtm campaign", "marketing"]
        pattern = "|".join(re.escape(k) for k in keywords)
        samples["_is_sales_team"] = desc_series.str.contains(pattern)
    else:
        samples["_is_sales_team"] = False

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

    # Post-process sales-team marked rows: override source/email and zero-out
    # financial/quantity fields (but keep total_shipping_cost).
    sales_idx = samples.index[samples["_is_sales_team"]]
    if len(sales_idx) > 0:
        # Mark source and email
        sample_rows.loc[sales_idx, "source"] = "sales_team"
        sample_rows.loc[sales_idx, "email"] = "SENT TO SALES TEAM"

        # Exclude from inventory counts: zero total_bars_sold
        if "total_bars_sold" in sample_rows.columns:
            sample_rows.loc[sales_idx, "total_bars_sold"] = 0

        # Zero-out financial columns to the right of quantity except shipping costs
        zero_cols = [
            "line_item_price",
            "subtotal",
            "discount",
            "shipping",
            "tax",
            "total",
            "bar_cogs",
        ]
        for col in zero_cols:
            if col in sample_rows.columns:
                sample_rows.loc[sales_idx, col] = 0

    # Drop helper column
    if "_is_sales_team" in samples.columns:
        samples = samples.drop(columns=["_is_sales_team"])

    # ---- FINAL MASTER LOG ----

    master = pd.concat([merged_rows, sample_rows], ignore_index=True)

    if output_path:
        out_path = Path(output_path)
        suffix = out_path.suffix.lower()
        # write CSV by default for .csv or unknown extensions
        if suffix in (".xlsx", ".xls"):
            try:
                master.to_excel(output_path, index=False)
                print(f"Master log Excel written to: {output_path}")
            except ImportError:
                raise ImportError(
                    "Writing Excel files requires 'openpyxl' (pip install openpyxl)."
                )
        else:
            master.to_csv(output_path, index=False)
            print(f"Master log CSV written to: {output_path}")

    return master


if __name__ == "__main__":
    # Update these paths as needed
    orders_file = "/Users/samskanse/desktop/orders_11-21_to_11-28.csv"
    # threepl_file = "Skye Performance 11.17.25 to 11.23.25.xlsx"
    # output_file = "/Users/samskanse/desktop/order_log_11-21_to_11-28.csv"

    # build_master_log(orders_file, threepl_file, output_file)
    orders_log_from_csv(
        orders_file,
        output_path="/Users/samskanse/desktop/orders_only_log_11-21_to_11-28.xlsx",
    )
