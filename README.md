# TODO
- Make logic for unfufilled orders (as orders fufilled long after submitted will needed to be included on current order log with 3pl to create period report)


# Skye-Financials-Automation

Sales Automation Task:

Goal: Output to a spreadsheet weekly to get our basic financial and product numbers
once we get our shipping cost (3PL Calibrate)


Formatting: 
output to a excel spreadsheet I create named “Skye_Period_Report_yryr_mm-dd_to_yryr_mm-dd”

Numbers needed:

Financials (per period):
Gross Revenue (Revenue + Shipping Charges)
Taxes Collected
COGS (of bars sold in period)
Total 3PL costs
Gross Profit
Gross Margin

Inventory
Note: 7 bars per box
Starting Inventory of period
Starting inventory at 3pl inception (before first sale on october 24th run #1)
15802 (maybe)
Boxes Sold
Bars Sold
Total Inventory Sold (Boxes + Bars)
# Skye-Financials-Automation

Automates weekly financial and product reporting by combining Shopify order data with 3PL (Calibrate) shipping/handling records. The pipeline builds a Master Log and a Financial & Inventory summary, then writes a two-tab Excel report named like `Skye_Period_Report_YYYY_MM-DD_to_YYYY_MM-DD`.

## Goal

Produce a clean, repeatable weekly report that includes:
- Financial metrics (revenue, taxes, COGS, 3PL costs, gross profit, margin)
- Inventory movement (boxes and bars sold, starting/ending inventory)

## Output filename

Standard output filename: `Skye_Period_Report_YYYY_MM-DD_to_YYYY_MM-DD.xlsx`

## Key assumptions and constants
- 7 bars per box
- Bar COGS is computed from a historical total (example: 39,891.91 / 15,848 ≈ 2.52 per bar)

## Inputs
- `Shopify` weekly CSV of orders
- `3PL` (Calibrate) spreadsheet with shipping, handling, packaging, and sample rows

## High-level Flow
1. Build a Master Log by merging Shopify and 3PL data into a single dataframe (one row per Shopify order, plus sample rows from 3PL where applicable).
2. Compute weekly financials and inventory summary from the Master Log.
3. Export a two-tab Excel workbook containing the Master Log and the Financial Summary.

## Where code lives
The implementation for each pipeline step is organized into modules under the `skyepipeline_files` package:

- `skyepipeline_files/MasterLogCreation.py`: builds the Master Log (merges Shopify orders with 3PL shipments and handles free samples / sales-team sendouts).
- `skyepipeline_files/WeeklySummaryCreator.py`: computes the Financial & Inventory summary DataFrame from the Master Log.
- `skyepipeline_files/BuildWeeklyWorkbook.py`: writes the final two-tab Excel workbook (`Master Log` + `Financial Summary`) from the pipeline DataFrames.
- `SkyePipeline.py`: top-level runner that orchestrates the steps (reads inputs, calls the three modules above, and writes the report).

## Master Log (per order row)
The Master Log contains three logical groups of fields:

- Orders details
    - `order_ID` (Shopify `Name`)
    - `order_date`
    - `email`
    - `box_or_bar` (determine from price: box if line item price > 20, bar if < 7)
    - `source` (channel)
    - `line_item_quantity`
    - `total_bars_sold` (if box: `line_item_quantity * 7`; if bar: `line_item_quantity * 1`)
    - `line_item_price`

- Order financials
    - `subtotal`
    - `discount`
    - `shipping` (collected)
    - `tax`
    - `total` (subtotal + shipping + tax)

- Costs
    - `bar_cogs` (per-bar COGS × bars sold for the row)
    - `total_shipping_cost` (3PL: Handling Fee + Total Shipping + Packaging)

### Free samples (3PL rows without `store order number`)
- Treated as separate Master Log rows. Fields populated from 3PL where available:
    - `email` = `FREE SAMPLES` (label for sample rows)
    - `line_item_quantity` = `Total Quantity` (from 3PL)
    - `box_or_bar` deduced from `(total price / quantity)`: box if > 20, bar if < 10
    - `total_shipping_cost` computed from 3PL columns (Handling Fee + Shipping + Packaging)
    - `tax` and `discount` copied from 3PL when present

### Sales-team / GTM sendouts
Some 3PL sample rows are internal or marketing sendouts (sales-team / GTM). These rows are detected in the 3PL data and handled specially so they don't affect sales/inventory metrics.

- Detection
    - If the 3PL row has no `Store Order Number` and the 3PL `Description` (or similar column) contains case-insensitive keywords such as `gtm`, `sales team`, `sales_team`, `gtm campaign`, or `marketing`, the row is flagged as a sales-team sendout.

- Handling rules
    - `source` is set to `sales_team` and `email` is set to `SENT TO SALES TEAM` for these rows.
    - Inventory impact is removed: `total_bars_sold` (and related inventory counts) is zeroed for sales-team rows so they do not count toward boxes/bars sold.
    - Financial fields to the right of quantity (unit price, subtotal, discount, shipping collected, tax, total, bar_cogs) are zeroed to avoid inflating revenue or COGS. The 3PL shipping cost is retained.
    - Sales-team rows are excluded when summing Boxes Sold / Bars Sold in the weekly summary.

Note: Detection relies on the presence of a Description-like column in the 3PL sheet; the keyword list can be adjusted later if needed.

## Important merge note
Match Shopify orders to 3PL rows by order ID: Shopify `Name` and 3PL `Store Order Number` (format `#1234`).

## Weekly Financials & Inventory Summary
The summary aggregates the Master Log for the period and reports these metrics:

- a. Gross Revenue (includes shipping collected)
    - Computed as `Sum(total)` (includes collected shipping and taxes as recorded)
- b. Taxes Collected
    - `Sum(tax)`
- c. COGS
    - `Sum(bar_cogs)`
- d. Shipping costs (3PL)
    - `Sum(total_shipping_cost)` + receiving (if present in 3PL) + payment processing fee (prompted at runtime, two decimals)
- e. Gross Profit
    - `Gross Revenue - COGS - Shipping Costs`
- f. Gross Margin
    - `Gross Profit / Gross Revenue`

- g. Starting Inventory (bars)
    - Prompted at runtime (whole number)
- h. Boxes Sold this week
    - Sum of `line_item_quantity` for rows where `box_or_bar == box`
- i. Bars Sold this week
    - Sum of `line_item_quantity` for `bar` rows
- j. Total Inventory Sold (bars)
    - Boxes converted to bars (`boxes * 7`) + bars sold
- k. Weekly Ending Inventory
    - `Starting Inventory - Total Inventory Sold`

### Summary layout for the report
- Column 1: Metric name (A–F)
- Column 2: Metric value
- Column 3: blank separator
- Column 4: Metric name (G–K)
- Column 5: Metric value

Example visual section header in the output workbook:

============== Cumulative Period Financials =====================

Revenue

 \+ Shipping collected

 —-------------------------

 Gross Revenue

 \+ Taxes (collected)

 \- COGS

 \- Total 3PL Costs (shipping, receiving, payment processing fee)

 —------------------------------------------------------

 Gross Profit

 Gross Margin

=============== Inventory / Units =======================
 
 Starting Inventory (bars)
 
 Boxes Sold This Period
 
 Bars Sold This Period (single bars)
 
 Total Inventory Sold (bars)
 
 Weekly Ending Inventory (bars)

> NOTE: Financial metric names may be adjusted for clarity later.

## Final export
The pipeline writes a two-tab Excel workbook:
- Tab 1: `Master Log` (detailed per-order rows)
- Tab 2: `Financial Summary` (metrics and inventory table)

## Future additions
- Get financial metrics by channel
- Use Shopify API for a more robust pipeline
- Add per-week sales rate (exclude samples and internal purchases)

## Open questions
- How many boxes were sent out to Ciaran/other sources before 3PL inception?
- Were those part of run 1 as well?

---
This README preserves the original requirements and layout while organizing the content into clear sections for easier reference and implementation.


