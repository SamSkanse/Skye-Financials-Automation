# Skye-Financials-Automation

Sales Automation Task:

Goal: Output to a spreadsheet weekly to get our basic financial and product numbers
once we get our shipping cost (3PL Calibrate)


Formatting: 
Excel Spreadsheet I create named “Week of xx/xx/xx Skye Financials Summary”
Which will be renamed replacing the Xs for that week

Numbers needed:

Financials
Gross Revenue (Includes our shipping charges)
Taxes Collected
COGS (per bar)
Shipping costs (3PL)
Gross Profit
Gross Margin

Bars Sold
Note: 7 bars per box
Starting Inventory
Starting inventory at 3pl inception (before first sale on october 24th run #1)
15802
Boxes Sold
Bars Sold
Total Inventory Sold (Boxes + Bars)
Weekly Ending Inventory

NOTE: Above will do categories, 1. Overall 2. By each channel 


Flow of operations:
IMMEDIATE TERM

Inputs: Shopify Weekly Csv of orders, 3pl charges spreadsheet, template spreadsheet

Ops:

1st: 
Take Shopify csv orders and 3pl spreadsheet and merge them into a single csv file in format below with only features I am outlining (only use things in parentheses around the features for references of rule or where to find the info from in current inputs



DO NOT INCLUDE CATEGORIES (Ex. Order Details) IN CSV, JUST THE FEATURES IN CURRENT ORDER


Important: To merge understand which 3PL row goes with what orders csv row, each with have an order ID (Name in orders csv and Store order number in 3PL spreadsheet) with format “#1234”




Tab for Master log for each row (order) includes this data below
Orders Details
Features: order_ID (order csv says name), order date, email, was it a box or bar sold (box = (line item price > 20), bar = (line item price < 5) ) , source, line item quantity, total bars sold (if box then line item quantity *7, if bar then line item quantity *1), line item price
Order Financials
Features: Subtotal, discount, shipping, tax, total (subtotal + shipping + tax)
Costs
Features: Bar COGS (= 39,891.91/15848 = 2.51…., if bar sold then 2.51…*line item quantity, if box sold then 2.51…*line item quantity*7), Total shipping cost (Handling Fee + Total Shipping + Packaging, from 3PL (Calibrate Shipping Charges))

NOTE: IF there is no order number in column “store order number” in 3PL sheet, then order was a free sample and features that should be filled in are below as such in the master log as another row entry (not necessarily in correct order below) (if I don’t give a feature for it leave it blank)

Email = FREE SAMPLE BOX, line item quantity = Total Quantity (column from 3PL sheet), was it a box or bar sold (box = ( total price /quantity > 20), bar = (total price/ quantity < 10), info from 3PL sheet), Total shipping cost (Handling Fee + Total Shipping + Packaging, from 3PL (Calibrate Shipping Charges)), tax (tax column), discount (discount column)

Output: csv file

2nd:
Take our master log and come up with our weekly financials and inventory summary

Formatting:
Cumulative Weekly Financials
Overall financials (features below a. b. … etc.)
Gross Revenue (Includes our shipping charges)
Sum of Total - Tax - shipping 
Taxes Collected
Sum of taxes
COGS (per bar)
Sum of COGS
Shipping costs (3PL)
Sum of total shipping costs + sum of receiving (a column in 3PL spreadsheet) + payment processing fee (ask user for it when code runs, 2 up to 2 decimal places)
Gross Profit
Gross Rev - COGS - Shipping Costs
Gross Margin
Gross Profit / Gross Revenue
Starting Inventory 
Ask for starting inventory when code is run, whole number
Boxes Sold this week
If order = box; box total += line item quantity
Bars Sold this week
If order = bar; bar total += line item quantity
Total Inventory Sold (Boxes + Bars)
Sum total bars sold (from master log csv)
Weekly Ending Inventory
Start inventory - total inventory sold


Column 1 Names of metrics A-F
Column 2 the corresponding numbers
Column 3 blank 
Column 4 names of metrics E-K
Column 5 Corresponding numbers

3rd:
Take our final two csv files and neatly outputs spreadsheet with 2 tabs
1 tab with master log of the week 
2 tab with Financials and inventory numbers of the week


Format of financials summary:
============== Cumulative Period Financials =====================
Revenue
+ Shipping collected
—-------------------------
Gross Revenue
+ Taxes
- COGS
- Total 3PL Costs (shipping, receiving, payment processing fee
—------------------------------------------------------
Gross Profit
Gross Margin
===============Inventory / Units=======================
Starting Inventory (bars)
Boxes Sold This Period
Bars Sold This Period (single bars
Total Inventory Sold (bars)
Weekly Ending Inventory (bars)




Table 2 Online Store Weekly Financials
How: (orders csv file → source → web)
(Later Can use shopify API for simplicity of automation)









GM by channel
NET profit by channel 
Current inventory
Sales rate/week (forecast over time to see when we would run out based on this current sales rate. This will allow us to strategize whether we want to try to sell product faster and explore other channels or whether we should conserve product)
Just want to get a sense for current inventory + the rate of boxes sold per week

