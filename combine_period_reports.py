# this file is to used to combine multiple period reports into one file
# Function:
# Will combine our master logs from each period into one master log
# Will combine our weekly summaries from each period into one weekly summary
# Will create a new excel file with the combined master log and weekly summary


# WORK IN PROGRESS - NOT YET FUNCTIONAL

import pandas as pd
import numpy as np
import re
from skyepipeline_files.MasterLogCreation import build_master_log
from skyepipeline_files.WeeklySummaryCreator import build_weekly_summary
from skyepipeline_files.BuildWeeklyWorkbook import build_weekly_workbook
import os
from datetime import datetime
import sys
from pathlib import Path
import tkinter as tk 
from tkinter import filedialog


def pick_report_files(title="Select period report files", filetypes=(('Excel files', ('*.xlsx', '*.xls')), ('All files', '*.*'))):
	"""Open a file picker allowing multiple selection and return list of paths or None if cancelled."""

	root = tk.Tk()
	root.withdraw()
	files = filedialog.askopenfilenames(title=title, filetypes=filetypes)
	root.destroy()
	if not files:
		return None
	return list(files)


def pick_output_directory(title="Select output folder for combined report"):
	"""Open a directory picker and return the selected path or None if cancelled."""
	root = tk.Tk()
	root.withdraw()
	
	# Default to Desktop
	desktop = Path.home() / "Desktop"
	initial_dir = str(desktop) if desktop.exists() else str(Path.home())
	
	dir_path = filedialog.askdirectory(title=title, initialdir=initial_dir)
	root.destroy()
	if not dir_path:
		return None
	return dir_path


def combine_master_logs(report_files, output_path=None, dedupe=False, primary_key=None):
	"""Combine 'Master Log' sheets from multiple period report Excel files.

	Args:
		report_files (iterable): list/tuple of Excel file paths.
		output_path (str or Path, optional): if provided, write combined sheet to this Excel file.
		dedupe (bool): if True and primary_key provided, drop duplicate rows by primary_key.
		primary_key (str or list, optional): column(s) to use for deduplication.

	Returns:
		pandas.DataFrame or None: combined DataFrame, or None if no master sheets found.
	"""
	if not report_files:
		raise ValueError("report_files must be a non-empty iterable of file paths")

	combined_frames = []
	for f in report_files:
		p = Path(f)
		if not p.exists():
			print(f"Warning: file not found, skipping: {p}")
			continue
		try:
			xls = pd.ExcelFile(p)
		except Exception as e:
			print(f"Warning: unable to read Excel file {p}: {e}")
			continue

		# Find a sheet that matches 'Master Log' (case-insensitive). Accept names
		# that equal 'master log' or contain both 'master' and 'log'.
		candidate_sheet = None
		for s in xls.sheet_names:
			s_norm = str(s).strip().lower()
			if s_norm == "master log" or ("master" in s_norm and "log" in s_norm):
				candidate_sheet = s
				break

		if candidate_sheet is None:
			print(f"No Master Log sheet found in {p.name}; skipping.")
			continue

		try:
			df = pd.read_excel(p, sheet_name=candidate_sheet)
		except Exception as e:
			print(f"Warning: failed to read sheet '{candidate_sheet}' in {p.name}: {e}")
			continue

		# Determine a period label for this source. Prefer a YYYY-MM-DD_to_YYYY-MM-DD
		# token appearing in the sheet title (candidate_sheet) or filename. If found,
		# format as MM/DD/YY-MM/DD/YY. Otherwise fall back to the filename.
		period_label = None
		# look in the sheet name first (candidate_sheet), then filename
		date_token = None
		for hay in (candidate_sheet, p.name):
			if not hay:
				continue
			m = re.search(r"(\d{4}-\d{2}-\d{2})\s*_?to\s*_?(\d{4}-\d{2}-\d{2})", str(hay), re.IGNORECASE)
			if m:
				date_token = (m.group(1), m.group(2))
				break

		if date_token:
			try:
				start_dt = datetime.strptime(date_token[0], "%Y-%m-%d")
				end_dt = datetime.strptime(date_token[1], "%Y-%m-%d")
				# format as MM/DD/YY-MM/DD/YY
				period_label = f"{start_dt.strftime('%m/%d/%y')}-{end_dt.strftime('%m/%d/%y')}"
			except Exception:
				period_label = None

		if period_label is None:
			period_label = p.name

		# annotate source period and append
		df["source_period_report"] = period_label
		combined_frames.append(df)

	if not combined_frames:
		print("No Master Log sheets found in any provided files.")
		return None

	combined = pd.concat(combined_frames, ignore_index=True, sort=False)

	if dedupe and primary_key is not None:
		combined = combined.drop_duplicates(subset=primary_key, keep="first")

	# Remove exclude_from_bars_sold column
	exclude_col = "exclude_from_bars_sold"
	matching_cols = [c for c in combined.columns if str(c).strip().lower() == exclude_col.lower()]
	if matching_cols:
		combined = combined.drop(columns=matching_cols)
	
	# Combine box_or_bar and box_or_bar_or_case columns
	box_or_bar_cols = [c for c in combined.columns if str(c).strip().lower() == "box_or_bar"]
	box_or_bar_or_case_cols = [c for c in combined.columns if str(c).strip().lower() == "box_or_bar_or_case"]
	
	if box_or_bar_cols and box_or_bar_or_case_cols:
		# Both columns exist - combine them (prefer box_or_bar_or_case, fallback to box_or_bar)
		combined[box_or_bar_or_case_cols[0]] = combined[box_or_bar_or_case_cols[0]].fillna(combined[box_or_bar_cols[0]])
		combined = combined.drop(columns=box_or_bar_cols)
	elif box_or_bar_cols and not box_or_bar_or_case_cols:
		# Only box_or_bar exists - rename it to box_or_bar_or_case
		combined = combined.rename(columns={box_or_bar_cols[0]: "box_or_bar_or_case"})
	# If only box_or_bar_or_case exists, keep it as is

	return combined


def combine_financial_summaries(report_files):
	"""Extract cumulative period financials from each report's Financial Summary sheet.

	Returns a DataFrame with one row per report and columns:
	['source_period_report','source_file','Revenue','Shipping_Collected','Gross_Revenue',
	 'Taxes_Collected','COGS_Total','Shipping_Costs_Total','Gross_Profit','Gross_Margin',
	 'rev_ship_match','profit_match']
	"""
	if not report_files:
		raise ValueError("report_files must be a non-empty iterable of file paths")

	# Aggregate totals across all provided files
	totals = {
		'Revenue': 0.0,
		'Shipping_Collected': 0.0,
		'Gross_Revenue': 0.0,
		'Taxes_Collected': 0.0,
		'COGS_Total': 0.0,
		'Shipping_Costs_Total': 0.0,
		'Gross_Profit': 0.0,
	}

	for f in report_files:
		p = Path(f)
		if not p.exists():
			print(f"Warning: file not found, skipping: {p}")
			continue
		try:
			xls = pd.ExcelFile(p)
		except Exception as e:
			print(f"Warning: unable to read Excel file {p}: {e}")
			continue

		# find candidate financial summary sheet
		candidate_sheet = None
		for s in xls.sheet_names:
			s_norm = str(s).strip().lower()
			if s_norm == "financial summary" or ("financial" in s_norm and "summary" in s_norm):
				candidate_sheet = s
				break

		# fallback: scan sheets for a 'Metric' column or presence of 'Gross Revenue' text
		if candidate_sheet is None:
			for s in xls.sheet_names:
				try:
					df_test = pd.read_excel(p, sheet_name=s, nrows=50)
					expl = [str(c).lower() for c in df_test.columns]
					# check for Metric column
					if any('metric' == c for c in expl) or any('metric' in c for c in expl):
						candidate_sheet = s
						break
					# or check first column for 'gross revenue'
					first_col = df_test.columns[0] if len(df_test.columns) > 0 else None
					if first_col is not None:
						if df_test[first_col].astype(str).str.lower().str.contains('gross revenue', na=False).any():
							candidate_sheet = s
							break
				except Exception:
					continue

		if candidate_sheet is None:
			print(f"No Financial Summary sheet found in {p.name}; skipping.")
			continue

		try:
			df = pd.read_excel(p, sheet_name=candidate_sheet)
		except Exception as e:
			print(f"Warning: failed to read sheet '{candidate_sheet}' in {p.name}: {e}")
			continue

		# normalize column names
		df.columns = [str(c).strip() for c in df.columns]

		# find metric and value columns
		metric_col = None
		value_col = None
		for c in df.columns:
			cl = str(c).lower()
			if 'metric' in cl:
				metric_col = c
			if 'value' in cl:
				value_col = c
		# fallback heuristics
		if metric_col is None:
			metric_col = df.columns[0]
		if value_col is None and len(df.columns) > 1:
			value_col = df.columns[1]

		# helper to parse numeric-like cells
		def parse_num(v):
			if pd.isna(v):
				return np.nan
			if isinstance(v, (int, float, np.number)):
				return float(v)
			s = str(v).strip()
			if s == '':
				return np.nan
			# parentheses for negatives
			neg = False
			if s.startswith('(') and s.endswith(')'):
				neg = True
				s2 = s.strip('()')
				# remove $ and commas
				s2 = s2.replace('$','').replace(',','').replace('\xa0','').strip()
			# percent
			if s2.endswith('%'):
				try:
					v = float(s2.strip('%'))/100.0
					return -v if neg else v
				except Exception:
					return np.nan
			try:
				v = float(s2)
				return -v if neg else v
			except Exception:
				return np.nan

		# build a mapping of cleaned metric -> numeric value
		metric_series = df[metric_col].astype(str).fillna('').tolist()
		value_series = df[value_col].tolist() if value_col in df.columns else [np.nan]*len(metric_series)
		mapping = {}
		for mtxt, v in zip(metric_series, value_series):
			mt = str(mtxt).strip()
			mt_clean = re.sub(r"^[\+\-]\s*", "", mt).strip().lower()
			mapping[mt_clean] = parse_num(v)

		# helper to fetch by keywords
		def find_metric(*keywords):
			for k, val in mapping.items():
				if all(kw in k for kw in keywords):
					return val
			return np.nan

		# extract desired fields
		revenue = find_metric('revenue') if 'revenue' in mapping else np.nan
		# revenue may be ambiguous (revenue vs gross revenue). Prefer exact 'revenue' key
		if 'revenue' in mapping:
			revenue = mapping.get('revenue', revenue)

		shipping = find_metric('shipping', 'collected')
		gross_revenue = find_metric('gross revenue')
		taxes = find_metric('taxes') if 'taxes collected' in mapping or any('tax' in k for k in mapping.keys()) else find_metric('tax')
		cogs = find_metric('cogs')
		threepl = None
		# look for '3pl' or 'total 3pl' or 'total 3pl costs' or 'total 3pl costs'
		for k in mapping.keys():
			if '3pl' in k or ('total' in k and '3pl' in k) or 'total 3pl' in k or 'total 3pl costs' in k:
				threepl = mapping.get(k)
				break
		if threepl is None:
			threepl = find_metric('3pl')

		gross_profit = find_metric('gross profit')

		# Recalculate gross margin as gross_profit / gross_revenue
		recalc_gross_margin = np.nan
		if not pd.isna(gross_revenue) and gross_revenue != 0:
			recalc_gross_margin = gross_profit / gross_revenue if not pd.isna(gross_profit) else np.nan

		# Checks
		tol = 0.01
		rev_ship_match = False
		if not pd.isna(revenue) and not pd.isna(shipping) and not pd.isna(gross_revenue):
			rev_ship_match = abs((revenue + shipping) - gross_revenue) <= tol

		profit_match = False
		if not pd.isna(gross_revenue) and not pd.isna(cogs) and not pd.isna(threepl) and not pd.isna(gross_profit):
			profit_match = abs(gross_revenue - cogs - threepl - gross_profit) <= tol

		# determine period_label same as combine_master_logs
		period_label = None
		date_token = None
		for hay in (candidate_sheet, p.name):
			if not hay:
				continue
			m = re.search(r"(\d{4}-\d{2}-\d{2})\s*_?to\s*_?(\d{4}-\d{2}-\d{2})", str(hay), re.IGNORECASE)
			if m:
				date_token = (m.group(1), m.group(2))
				break
		if date_token:
			try:
				start_dt = datetime.strptime(date_token[0], "%Y-%m-%d")
				end_dt = datetime.strptime(date_token[1], "%Y-%m-%d")
				period_label = f"{start_dt.strftime('%m/%d/%y')}-{end_dt.strftime('%m/%d/%y')}"
			except Exception:
				period_label = p.name
		else:
			period_label = p.name

		# Add numeric values to the running totals (treat NaN as 0)
		def add_tot(key, val):
			try:
				totals[key] += 0.0 if pd.isna(val) else float(val)
			except Exception:
				pass

		add_tot('Revenue', revenue)
		add_tot('Shipping_Collected', shipping)
		add_tot('Gross_Revenue', gross_revenue)
		add_tot('Taxes_Collected', taxes)
		add_tot('COGS_Total', cogs)
		add_tot('Shipping_Costs_Total', threepl)
		add_tot('Gross_Profit', gross_profit)

	# If we never added any totals, nothing to return
	if all(v == 0 for v in totals.values()):
		print("No financial summaries found in any provided files or all values zero.")
		return None

	# Recalculate gross margin using aggregated totals
	gross_margin = np.nan
	if totals['Gross_Revenue'] != 0:
		try:
			gross_margin = totals['Gross_Profit'] / totals['Gross_Revenue']
		except Exception:
			gross_margin = np.nan

	# Build a pretty summary DataFrame similar to BuildWeeklyWorkbook
	rows_pretty = []
	rows_pretty.append(["============== Cumulative Period Financials =====================", ""]) 
	rows_pretty.append(["Revenue", totals['Revenue']])
	rows_pretty.append(["+ Shipping collected", totals['Shipping_Collected']])
	rows_pretty.append(["------------------------------------------------------", ""])
	rows_pretty.append(["Gross Revenue", totals['Gross_Revenue']])
	rows_pretty.append(["+ Taxes Collected", totals['Taxes_Collected']])
	# subtractive values shown as negatives
	try:
		cogs_value = -abs(float(totals['COGS_Total']))
	except Exception:
		cogs_value = -abs(totals['COGS_Total']) if totals['COGS_Total'] is not None else np.nan
	try:
		threepl_value = -abs(float(totals['Shipping_Costs_Total']))
	except Exception:
		threepl_value = -abs(totals['Shipping_Costs_Total']) if totals['Shipping_Costs_Total'] is not None else np.nan

	rows_pretty.append(["- COGS", cogs_value])
	rows_pretty.append(["- Total 3PL Costs (shipping, receiving, payment processing fee)", threepl_value, ""])
	rows_pretty.append(["------------------------------------------------------", ""])
	rows_pretty.append(["Gross Profit", totals['Gross_Profit']])
	rows_pretty.append(["Gross Margin", gross_margin if not pd.isna(gross_margin) else "N/A"]) 

	summary_pretty = pd.DataFrame(rows_pretty, columns=["Metric", "Value", "Note"])

	# Also provide a numeric one-row summary
	numeric_summary = pd.DataFrame([{
		'Revenue': totals['Revenue'],
		'Shipping_Collected': totals['Shipping_Collected'],
		'Gross_Revenue': totals['Gross_Revenue'],
		'Taxes_Collected': totals['Taxes_Collected'],
		'COGS_Total': totals['COGS_Total'],
		'Shipping_Costs_Total': totals['Shipping_Costs_Total'],
		'Gross_Profit': totals['Gross_Profit'],
		'Gross_Margin': gross_margin,
	}])

	return numeric_summary, summary_pretty


def combine_inventory_summaries(report_files):
	"""Extract cumulative period inventory data from each report's Financial Summary sheet.
	
	Returns a DataFrame with inventory-related rows formatted for display:
	['Starting_Inventory_Bars', 'Cases_Sold', 'Boxes_Sold', 'Bars_Sold', 
	 'Total_Inventory_Sold', 'Weekly_Ending_Inventory_Bars']
	"""
	if not report_files:
		raise ValueError("report_files must be a non-empty iterable of file paths")
	
	# Store inventory data with period dates for sorting
	inventory_data = []
	
	for f in report_files:
		p = Path(f)
		if not p.exists():
			continue
		
		try:
			xls = pd.ExcelFile(p)
		except Exception:
			continue
		
		# Find Financial Summary sheet
		candidate_sheet = None
		for s in xls.sheet_names:
			s_norm = str(s).strip().lower()
			if s_norm == "financial summary" or ("financial" in s_norm and "summary" in s_norm):
				candidate_sheet = s
				break
		
		if candidate_sheet is None:
			continue
		
		try:
			df = pd.read_excel(p, sheet_name=candidate_sheet)
		except Exception:
			continue
		
		# Normalize columns
		df.columns = [str(c).strip() for c in df.columns]
		
		# Find metric and value columns
		metric_col = df.columns[0] if len(df.columns) > 0 else None
		value_col = df.columns[1] if len(df.columns) > 1 else None
		
		if metric_col is None or value_col is None:
			continue
		
		# Helper to parse numeric values
		def parse_num(v):
			if pd.isna(v):
				return np.nan
			if isinstance(v, (int, float, np.number)):
				return float(v)
			s = str(v).strip()
			if s == '':
				return np.nan
			# Handle negatives in parentheses
			neg = False
			if s.startswith('(') and s.endswith(')'):
				neg = True
				s = s.strip('()')
			s = s.replace('$','').replace(',','').replace('\xa0','').strip()
			try:
				v = float(s)
				return -v if neg else v
			except Exception:
				return np.nan
		
		# Build mapping
		metric_series = df[metric_col].astype(str).fillna('').tolist()
		value_series = df[value_col].tolist()
		mapping = {}
		for mtxt, v in zip(metric_series, value_series):
			mt = str(mtxt).strip()
			mt_clean = re.sub(r"^[\+\-]\s*", "", mt).strip().lower()
			mapping[mt_clean] = parse_num(v)
		
		# Helper to find metrics by keywords
		def find_metric(*keywords):
			for k, val in mapping.items():
				if all(kw in k for kw in keywords):
					return val
			return np.nan
		
		# Extract inventory metrics
		starting_inv = find_metric('starting inventory')
		cases_sold = find_metric('cases sold', 'this period') if find_metric('cases sold', 'this period') == find_metric('cases sold', 'this period') else find_metric('cases sold')
		boxes_sold = find_metric('boxes sold', 'this period') if find_metric('boxes sold', 'this period') == find_metric('boxes sold', 'this period') else find_metric('boxes sold')
		bars_sold = find_metric('single bars sold', 'this period') if find_metric('single bars sold', 'this period') == find_metric('single bars sold', 'this period') else find_metric('bars sold')
		
		case_bars = find_metric('case bars sold')
		box_bars = find_metric('box bars sold')
		single_bars = find_metric('single bars sold') if 'single bars sold' in ' '.join(mapping.keys()) else bars_sold
		
		total_inv_sold = find_metric('total inventory sold')
		ending_inv = find_metric('ending inventory')
		
		# Extract period date for sorting (earliest to latest)
		period_date = None
		date_token = None
		for hay in (candidate_sheet, p.name):
			if not hay:
				continue
			m = re.search(r"(\d{4}-\d{2}-\d{2})\s*_?to\s*_?(\d{4}-\d{2}-\d{2})", str(hay), re.IGNORECASE)
			if m:
				date_token = (m.group(1), m.group(2))
				break
		
		if date_token:
			try:
				start_dt = datetime.strptime(date_token[0], "%Y-%m-%d")
				period_date = start_dt
			except Exception:
				period_date = None
		
		inventory_data.append({
			'period_date': period_date,
			'starting_inventory': starting_inv,
			'cases_sold': cases_sold,
			'boxes_sold': boxes_sold,
			'bars_sold': bars_sold,
			'case_bars': case_bars,
			'box_bars': box_bars,
			'single_bars': single_bars,
			'total_inv_sold': total_inv_sold,
			'ending_inv': ending_inv,
		})
	
	if not inventory_data:
		print("No inventory summaries found in any provided files.")
		return None
	
	# Sort by period_date to get earliest first
	inventory_data.sort(key=lambda x: x['period_date'] if x['period_date'] else datetime.max)
	
	# Starting inventory from earliest period
	starting_inventory = 0
	for item in inventory_data:
		if not pd.isna(item['starting_inventory']):
			starting_inventory = int(item['starting_inventory'])
			break
	
	# Sum all other metrics
	totals = {
		'cases_sold': 0,
		'boxes_sold': 0,
		'bars_sold': 0,
		'case_bars': 0,
		'box_bars': 0,
		'single_bars': 0,
		'total_inv_sold': 0,
	}
	
	for item in inventory_data:
		for key in totals.keys():
			if not pd.isna(item[key]):
				totals[key] += int(item[key])
	
	# Ending inventory from latest period or calculate
	ending_inventory = 0
	for item in reversed(inventory_data):
		if not pd.isna(item['ending_inv']):
			ending_inventory = int(item['ending_inv'])
			break
	
	# If no ending inventory found, calculate it
	if ending_inventory == 0:
		ending_inventory = starting_inventory - totals['total_inv_sold']
	
	# ======================================================
	#           VALIDATION CHECKS
	# ======================================================
	
	# Check 1: Bars from cases/boxes/singles should equal total inventory sold
	calculated_total_bars = totals['case_bars'] + totals['box_bars'] + totals['single_bars']
	bars_match = abs(calculated_total_bars - totals['total_inv_sold']) < 1  # Allow 1 bar tolerance
	bars_match_note = ""
	if not bars_match:
		bars_match_note = f"WARNING: Bars sum mismatch! Calculated={calculated_total_bars}, Total={totals['total_inv_sold']}, Diff={calculated_total_bars - totals['total_inv_sold']}"
		print(f"⚠️  {bars_match_note}")
	
	# Check 2: Starting - Total Sold should equal Ending
	calculated_ending = starting_inventory - totals['total_inv_sold']
	inventory_balance_match = abs(calculated_ending - ending_inventory) < 1  # Allow 1 bar tolerance
	inventory_balance_note = ""
	if not inventory_balance_match:
		inventory_balance_note = f"WARNING: Inventory balance mismatch! Starting({starting_inventory}) - Sold({totals['total_inv_sold']}) = {calculated_ending}, but Ending={ending_inventory}, Diff={calculated_ending - ending_inventory}"
		print(f"⚠️  {inventory_balance_note}")
	
	# Build formatted rows
	rows = []
	rows.append(["", "", ""])
	rows.append(["===============Inventory / Units=======================", "", ""])
	rows.append(["Starting Inventory (bars)", starting_inventory, ""])
	rows.append(["", "", ""])
	rows.append(["Cases Sold (units)", totals['cases_sold'], ""])
	rows.append(["Boxes Sold (units)", totals['boxes_sold'], ""])
	rows.append(["Single Bars Sold (units)", totals['bars_sold'], ""])
	rows.append(["------------------------------------------------------", "", ""])
	rows.append(["Bars from Cases (168 each)", totals['case_bars'], ""])
	rows.append(["+ Bars from Boxes (7 each)", totals['box_bars'], ""])
	rows.append(["+ Bars from Singles (1 each)", totals['single_bars'], ""])
	rows.append(["------------------------------------------------------", "", ""])
	rows.append(["Total Inventory Sold (bars)", totals['total_inv_sold'], bars_match_note])
	rows.append(["", "", ""])
	
	try:
		total_inv_sold_value = -abs(int(totals['total_inv_sold']))
	except Exception:
		total_inv_sold_value = -abs(totals['total_inv_sold'])
	
	rows.append(["Starting Inventory (bars)", starting_inventory, ""])
	rows.append(["- Total Inventory Sold (bars)", total_inv_sold_value, ""])
	rows.append(["------------------------------------------------------", "", ""])
	rows.append(["Ending Inventory (bars)", ending_inventory, inventory_balance_note])
	
	summary = pd.DataFrame(rows, columns=["Metric", "Value", "Note"])
	return summary


def combine_pos_summary(report_files):
	"""Extract cumulative period POS/remaining inventory data from each report's Financial Summary sheet.
	
	Returns a DataFrame with POS and remaining inventory rows formatted for display:
	['Total_POS_Bars_Sent', 'Bars_to_be_Sold_POS', 'Single_Bars_Sold', 
	 'Bars_Outstanding_POS', 'Bars_Left_at_3PL']
	"""
	if not report_files:
		raise ValueError("report_files must be a non-empty iterable of file paths")
	
	# Store POS data with period dates for sorting
	pos_data = []
	
	for f in report_files:
		p = Path(f)
		if not p.exists():
			continue
		
		try:
			xls = pd.ExcelFile(p)
		except Exception:
			continue
		
		# Find Financial Summary sheet
		candidate_sheet = None
		for s in xls.sheet_names:
			s_norm = str(s).strip().lower()
			if s_norm == "financial summary" or ("financial" in s_norm and "summary" in s_norm):
				candidate_sheet = s
				break
		
		if candidate_sheet is None:
			continue
		
		try:
			df = pd.read_excel(p, sheet_name=candidate_sheet)
		except Exception:
			continue
		
		# Normalize columns
		df.columns = [str(c).strip() for c in df.columns]
		
		# Find metric and value columns
		metric_col = df.columns[0] if len(df.columns) > 0 else None
		value_col = df.columns[1] if len(df.columns) > 1 else None
		
		if metric_col is None or value_col is None:
			continue
		
		# Helper to parse numeric values
		def parse_num(v):
			if pd.isna(v):
				return np.nan
			if isinstance(v, (int, float, np.number)):
				return float(v)
			s = str(v).strip()
			if s == '':
				return np.nan
			# Handle negatives in parentheses
			neg = False
			if s.startswith('(') and s.endswith(')'):
				neg = True
				s = s.strip('()')
			s = s.replace('$','').replace(',','').replace('\xa0','').strip()
			try:
				v = float(s)
				return -v if neg else v
			except Exception:
				return np.nan
		
		# Build mapping
		metric_series = df[metric_col].astype(str).fillna('').tolist()
		value_series = df[value_col].tolist()
		mapping = {}
		for mtxt, v in zip(metric_series, value_series):
			mt = str(mtxt).strip()
			mt_clean = re.sub(r"^[\+\-]\s*", "", mt).strip().lower()
			mapping[mt_clean] = parse_num(v)
		
		# Helper to find metrics by keywords
		def find_metric(*keywords):
			for k, val in mapping.items():
				if all(kw in k for kw in keywords):
					return val
			return np.nan
		
		# Extract POS metrics
		total_pos_bars = find_metric('total pos bars', 'sales members')
		if pd.isna(total_pos_bars):
			total_pos_bars = find_metric('total pos bars')
		
		single_bars_sold = find_metric('single bars sold')
		if pd.isna(single_bars_sold):
			single_bars_sold = find_metric('bars sold', 'single')
		
		bars_outstanding = find_metric('bars outstanding')
		if pd.isna(bars_outstanding):
			bars_outstanding = find_metric('outstanding', 'pos')
		
		ending_inventory = find_metric('ending inventory')
		bars_left_3pl = find_metric('bars left at 3pl')
		if pd.isna(bars_left_3pl):
			bars_left_3pl = find_metric('bars left', '3pl')
		
		# Extract period date for sorting (earliest to latest)
		period_date = None
		date_token = None
		for hay in (candidate_sheet, p.name):
			if not hay:
				continue
			m = re.search(r"(\d{4}-\d{2}-\d{2})\s*_?to\s*_?(\d{4}-\d{2}-\d{2})", str(hay), re.IGNORECASE)
			if m:
				date_token = (m.group(1), m.group(2))
				break
		
		if date_token:
			try:
				start_dt = datetime.strptime(date_token[0], "%Y-%m-%d")
				period_date = start_dt
			except Exception:
				period_date = None
		
		pos_data.append({
			'period_date': period_date,
			'total_pos_bars': total_pos_bars,
			'single_bars_sold': single_bars_sold,
			'bars_outstanding': bars_outstanding,
			'ending_inventory': ending_inventory,
			'bars_left_3pl': bars_left_3pl,
		})
	
	if not pos_data:
		print("No POS summaries found in any provided files.")
		return None
	
	# Sort by period_date to get earliest first, latest last
	pos_data.sort(key=lambda x: x['period_date'] if x['period_date'] else datetime.max)
	
	# Most recent values (from latest period)
	latest = pos_data[-1]
	total_pos_bars_given = 0
	if not pd.isna(latest['total_pos_bars']):
		total_pos_bars_given = int(latest['total_pos_bars'])
	
	bars_outstanding_latest = 0
	if not pd.isna(latest['bars_outstanding']):
		bars_outstanding_latest = int(latest['bars_outstanding'])
	
	ending_inventory_latest = 0
	if not pd.isna(latest['ending_inventory']):
		ending_inventory_latest = int(latest['ending_inventory'])
	
	bars_left_3pl_latest = 0
	if not pd.isna(latest['bars_left_3pl']):
		bars_left_3pl_latest = int(latest['bars_left_3pl'])
	
	# Sum single bars sold across all periods
	total_single_bars_sold = 0
	for item in pos_data:
		if not pd.isna(item['single_bars_sold']):
			total_single_bars_sold += int(item['single_bars_sold'])
	
	# ======================================================
	#           VALIDATION CHECKS
	# ======================================================
	
	# Check 1: Total bars given - single bars sold should equal remaining POS bars
	calculated_pos_remaining = total_pos_bars_given - total_single_bars_sold
	pos_balance_match = abs(calculated_pos_remaining - bars_outstanding_latest) < 1  # Allow 1 bar tolerance
	pos_balance_note = ""
	if not pos_balance_match:
		pos_balance_note = f"WARNING: POS balance mismatch! Given({total_pos_bars_given}) - Sold({total_single_bars_sold}) = {calculated_pos_remaining}, but Remaining={bars_outstanding_latest}, Diff={calculated_pos_remaining - bars_outstanding_latest}"
		print(f"⚠️  {pos_balance_note}")
	
	# Check 2: Ending inventory - bars given to sales team should equal bars left at 3PL
	calculated_3pl_bars = ending_inventory_latest - total_pos_bars_given
	warehouse_match = abs(calculated_3pl_bars - bars_left_3pl_latest) < 1  # Allow 1 bar tolerance
	warehouse_note = ""
	if not warehouse_match:
		warehouse_note = f"WARNING: Warehouse balance mismatch! Ending({ending_inventory_latest}) - Given({total_pos_bars_given}) = {calculated_3pl_bars}, but 3PL={bars_left_3pl_latest}, Diff={calculated_3pl_bars - bars_left_3pl_latest}"
		print(f"⚠️  {warehouse_note}")
	
	# Build formatted rows similar to BuildWeeklyWorkbook
	rows = []
	rows.append(["", "", ""])
	rows.append(["=============== POS / Sales Team Inventory ============", "", ""])
	rows.append(["POS Bars Sent to Sales Team", total_pos_bars_given, "Cumulative total"])
	rows.append(["Single Bars Sold by Sales Team", total_single_bars_sold, "Cumulative total"])
	rows.append(["------------------------------------------------------", "", ""])
	rows.append(["POS Bars Remaining", bars_outstanding_latest, pos_balance_note])
	rows.append(["", "", ""])
	rows.append(["", "", ""])
	rows.append(["=============== Warehouse Inventory ====================", "", ""])
	rows.append(["Ending Inventory (bars)", ending_inventory_latest, ""])
	
	# Show total POS bars given as subtractive
	try:
		total_pos_bars_value = -abs(int(total_pos_bars_given))
	except Exception:
		total_pos_bars_value = -abs(total_pos_bars_given)
	
	rows.append(["- POS Bars Sent to Sales Team", total_pos_bars_value, ""])
	rows.append(["------------------------------------------------------", "", ""])
	
	# Combine negative bars warning with warehouse mismatch warning
	pos_note = ""
	if bars_left_3pl_latest < 0:
		pos_note = f"WARNING: Negative bars at 3PL: {bars_left_3pl_latest}"
		print(f"⚠️  {pos_note}")
	
	# If there's both a warehouse note and a negative note, combine them
	final_3pl_note = warehouse_note
	if pos_note and warehouse_note:
		final_3pl_note = f"{warehouse_note} | {pos_note}"
	elif pos_note:
		final_3pl_note = pos_note
	
	rows.append(["Bars Remaining at 3PL", bars_left_3pl_latest, final_3pl_note])
	
	summary = pd.DataFrame(rows, columns=["Metric", "Value", "Note"])
	return summary


def combine_period_reports(report_files, output_path):
	"""Orchestrate combining all period reports into a single Excel workbook.
	
	Combines Master Logs and Financial Summaries (financials + inventory + POS)
	from multiple period report Excel files into a single output file.
	
	Args:
		report_files (iterable): list/tuple of Excel file paths.
		output_path (str or Path): path to write the combined Excel workbook.
	
	Returns:
		bool: True if successful, False otherwise.
	"""
	if not report_files:
		raise ValueError("report_files must be a non-empty iterable of file paths")
	
	print("Combining period reports...")
	
	# ---- COMBINE MASTER LOGS ----
	print("\n1. Combining Master Logs...")
	master_combined = combine_master_logs(report_files)
	if master_combined is None:
		print("Warning: No master logs found. Continuing with summaries only.")
		master_combined = pd.DataFrame()
	else:
		print(f"   ✓ Combined {len(master_combined)} rows from Master Logs")
	
	# ---- COMBINE FINANCIAL SUMMARIES ----
	print("\n2. Combining Financial Summaries...")
	fin_result = combine_financial_summaries(report_files)
	if fin_result is None:
		print("Warning: No financial summaries found. Continuing with other sections.")
		fin_summary = pd.DataFrame()
	else:
		_, fin_summary = fin_result
		print(f"   ✓ Combined Financial Summary ({len(fin_summary)} rows)")
	
	# ---- COMBINE INVENTORY SUMMARIES ----
	print("\n3. Combining Inventory Summaries...")
	inv_summary = combine_inventory_summaries(report_files)
	if inv_summary is None:
		print("Warning: No inventory summaries found. Continuing with other sections.")
		inv_summary = pd.DataFrame()
	else:
		print(f"   ✓ Combined Inventory Summary ({len(inv_summary)} rows)")
	
	# ---- COMBINE POS SUMMARIES ----
	print("\n4. Combining POS Summaries...")
	pos_summary = combine_pos_summary(report_files)
	if pos_summary is None:
		print("Warning: No POS summaries found. Continuing with other sections.")
		pos_summary = pd.DataFrame()
	else:
		print(f"   ✓ Combined POS Summary ({len(pos_summary)} rows)")
	
	# ---- CONCATENATE ALL SUMMARY SECTIONS ----
	print("\n5. Merging all summary sections...")
	summary_frames = []
	if not fin_summary.empty:
		summary_frames.append(fin_summary)
	if not inv_summary.empty:
		summary_frames.append(inv_summary)
	if not pos_summary.empty:
		summary_frames.append(pos_summary)
	
	if summary_frames:
		combined_summary = pd.concat(summary_frames, ignore_index=True, sort=False)
		print(f"   ✓ Merged summaries ({len(combined_summary)} total rows)")
	else:
		print("Warning: No summary sections available.")
		combined_summary = pd.DataFrame()
	
	# ---- WRITE TO EXCEL ----
	print(f"\n6. Writing to Excel: {output_path}")
	try:
		with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
			# Sheet 1: Master Log (if available)
			if not master_combined.empty:
				master_combined.to_excel(writer, sheet_name="Master Log", index=False)
				print(f"   ✓ Master Log sheet written ({len(master_combined)} rows)")
			
			# Sheet 2: Financial Summary (combined sections)
			if not combined_summary.empty:
				summary_escaped = combined_summary.copy()
				def _escape_cell(val):
					if isinstance(val, str) and val and val[0] in ('=', '+', '-'):
						return "'" + val
					return val
				if 'Metric' in summary_escaped.columns:
					summary_escaped['Metric'] = summary_escaped['Metric'].apply(_escape_cell)
				if 'Note' in summary_escaped.columns:
					summary_escaped['Note'] = summary_escaped['Note'].apply(_escape_cell)
				summary_escaped.to_excel(writer, sheet_name="Financial Summary", index=False)
				print(f"   ✓ Financial Summary sheet written ({len(combined_summary)} rows)")
				
				# Apply accounting format to numeric cells in Value column (column B)
				# BUT ONLY for the Financial Summary section (not Inventory or POS sections)
				wb = writer.book
				ws = wb["Financial Summary"]
				accounting_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"_);_(@_)'
				percentage_format = '0.00%'
				
				# Find the Value column (should be column B, index 2)
				value_col_idx = None
				metric_col_idx = None
				for idx, col in enumerate(ws[1], start=1):
					if col.value == "Value":
						value_col_idx = idx
					if col.value == "Metric":
						metric_col_idx = idx
				
				if value_col_idx and metric_col_idx:
					# Only apply to Financial Summary rows (first section)
					# Row 1 is header, financial summary starts at row 2
					fin_summary_end_row = 1 + len(fin_summary) if not fin_summary.empty else 1
					
					# Apply accounting format only to financial summary rows (except Gross Margin)
					for row in range(2, fin_summary_end_row + 1):
						metric_cell = ws.cell(row=row, column=metric_col_idx)
						value_cell = ws.cell(row=row, column=value_col_idx)
						
						# Only apply format if cell contains a number
						if isinstance(value_cell.value, (int, float, np.number)) and not pd.isna(value_cell.value):
							# Check if this is the Gross Margin row
							if metric_cell.value and 'gross margin' in str(metric_cell.value).lower():
								value_cell.number_format = percentage_format
							else:
								value_cell.number_format = accounting_format
			
			# Auto-size columns
			wb = writer.book
			for sheet_name in wb.sheetnames:
				ws = wb[sheet_name]
				for col in ws.columns:
					max_length = 0
					col_letter = col[0].column_letter
					for cell in col:
						try:
							cell_value = "" if cell.value is None else str(cell.value)
							if len(cell_value) > max_length:
								max_length = len(cell_value)
						except Exception:
							pass
					ws.column_dimensions[col_letter].width = min(max_length + 2, 50)
		
		print(f"\n✓ Successfully wrote combined report to: {output_path}")
		return True
	
	except Exception as e:
		print(f"\n✗ Failed to write Excel file: {e}")
		return False

if __name__ == "__main__":
	# Runner: open picker and print chosen files (or accept CLI args)
	args = sys.argv[1:]
	if args:
		chosen = args
	else:
		chosen = pick_report_files()

	if not chosen:
		print("No files selected.")
		sys.exit(0)

	# Prompt for output directory
	output_dir = pick_output_directory()
	if not output_dir:
		print("No output folder selected. Exiting.")
		sys.exit(0)

	print(f"Selected {len(chosen)} file(s):")
	for p in chosen:
		print(f" - {p}")

	print(f"Output folder: {output_dir}")

	# Menu for testing options
	print("\n" + "="*60)
	print("Select test option:")
	print("1. Test Master Log")
	print("2. Test Financial Summary")
	print("3. Test Inventory Summary")
	print("4. Test POS Summary")
	print("5. Test Full Combined Report (all sections)")
	print("="*60)
	
	choice = input("Enter choice (1-5): ").strip()

	if choice == "1":
		# --- Test: combine master logs ---
		print("\nCombining Master Log sheets from selected files...")
		master_result = combine_master_logs(chosen)
		if master_result is None:
			print("No master logs were combined. Exiting.")
			sys.exit(0)

		out_file = Path(output_dir) / "combined_master_log.xlsx"
		try:
			with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
				master_result.to_excel(writer, sheet_name="Master Log", index=False)
			print(f"Combined master log written to: {out_file}")
		except Exception as e:
			print(f"Failed to write combined master log to {out_file}: {e}")
			print("You can still inspect 'master_result' if running interactively.")

	elif choice == "2":
		# --- Test: combine financial summaries ---
		print("\nCombining Financial Summary sheets from selected files...")
		fin_result = combine_financial_summaries(chosen)
		if fin_result is None:
			print("No financial summaries were combined. Exiting.")
			sys.exit(0)

		_, fin_summary = fin_result
		out_file = Path(output_dir) / "combined_financial_summary.xlsx"
		try:
			with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
				summary_escaped = fin_summary.copy()
				def _escape_cell(val):
					if isinstance(val, str) and val and val[0] in ('=', '+', '-'):
						return "'" + val
					return val
				if 'Metric' in summary_escaped.columns:
					summary_escaped['Metric'] = summary_escaped['Metric'].apply(_escape_cell)
				if 'Note' in summary_escaped.columns:
					summary_escaped['Note'] = summary_escaped['Note'].apply(_escape_cell)
				summary_escaped.to_excel(writer, sheet_name="Financial Summary", index=False)
			print(f"Combined financial summary written to: {out_file}")
		except Exception as e:
			print(f"Failed to write combined financial summary to {out_file}: {e}")
			print("You can still inspect 'fin_summary' if running interactively.")

	elif choice == "3":
		# --- Test: combine inventory summaries ---
		print("\nCombining Inventory Summary sheets from selected files...")
		inv_result = combine_inventory_summaries(chosen)
		if inv_result is None:
			print("No inventory summaries were combined. Exiting.")
			sys.exit(0)

		out_file = Path(output_dir) / "combined_inventory_summary.xlsx"
		try:
			with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
				summary_escaped = inv_result.copy()
				def _escape_cell(val):
					if isinstance(val, str) and val and val[0] in ('=', '+', '-'):
						return "'" + val
					return val
				if 'Metric' in summary_escaped.columns:
					summary_escaped['Metric'] = summary_escaped['Metric'].apply(_escape_cell)
				if 'Note' in summary_escaped.columns:
					summary_escaped['Note'] = summary_escaped['Note'].apply(_escape_cell)
				summary_escaped.to_excel(writer, sheet_name="Inventory Summary", index=False)
			print(f"Combined inventory summary written to: {out_file}")
		except Exception as e:
			print(f"Failed to write combined inventory summary to {out_file}: {e}")
			print("You can still inspect 'inv_result' if running interactively.")

	elif choice == "4":
		# --- Test: combine POS summaries ---
		print("\nCombining POS Summary sheets from selected files...")
		pos_result = combine_pos_summary(chosen)
		if pos_result is None:
			print("No POS summaries were combined. Exiting.")
			sys.exit(0)

		out_file = Path(output_dir) / "combined_pos_summary.xlsx"
		try:
			with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
				summary_escaped = pos_result.copy()
				def _escape_cell(val):
					if isinstance(val, str) and val and val[0] in ('=', '+', '-'):
						return "'" + val
					return val
				if 'Metric' in summary_escaped.columns:
					summary_escaped['Metric'] = summary_escaped['Metric'].apply(_escape_cell)
				if 'Note' in summary_escaped.columns:
					summary_escaped['Note'] = summary_escaped['Note'].apply(_escape_cell)
				summary_escaped.to_excel(writer, sheet_name="POS Summary", index=False)
			print(f"Combined POS summary written to: {out_file}")
		except Exception as e:
			print(f"Failed to write combined POS summary to {out_file}: {e}")
			print("You can still inspect 'pos_result' if running interactively.")

	elif choice == "5":
		# --- Test: full combined report ---
		print("\nCombining full period report (all sections)...")
		out_file = Path(output_dir) / "combined_period_report.xlsx"
		success = combine_period_reports(chosen, out_file)
		if success:
			print(f"\nFull combined report successfully created: {out_file}")
		else:
			print("\nFailed to create full combined report.")
			sys.exit(1)

	else:
		print(f"\nInvalid choice: {choice}. Please run again and select 1-5.")
		sys.exit(1)

