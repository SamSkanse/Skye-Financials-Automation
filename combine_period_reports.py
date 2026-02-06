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
	dir_path = filedialog.askdirectory(title=title)
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

	return combined


def combine_weekly_summaries(report_files):
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


def combine_period_reports(report_files, output_path):
	pass

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

	# --- Test: combine weekly summaries and write to Excel in the chosen output folder ---
	print("\nCombining Financial Summary sheets from selected files...")
	result = combine_weekly_summaries(chosen)
	if result is None:
		print("No weekly summaries were combined. Exiting.")
		sys.exit(0)

	numeric_summary, summary_pretty = result

	out_file = Path(output_dir) / "combined_weekly_summary.xlsx"
	try:
		with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
			# Escape leading characters that Excel may interpret as formulas
			summary_escaped = summary_pretty.copy()
			def _escape_cell(val):
				if isinstance(val, str) and val and val[0] in ('=', '+', '-'):
					return "'" + val
				return val
			# Escape Metric and Note columns if present
			if 'Metric' in summary_escaped.columns:
				summary_escaped['Metric'] = summary_escaped['Metric'].apply(_escape_cell)
			if 'Note' in summary_escaped.columns:
				summary_escaped['Note'] = summary_escaped['Note'].apply(_escape_cell)
			# write pretty summary similar to BuildWeeklyWorkbook
			summary_escaped.to_excel(writer, sheet_name="Financial Summary", index=False)
		print(f"Combined weekly summary written to: {out_file}")
	except Exception as e:
		print(f"Failed to write combined weekly summary to {out_file}: {e}")
		print("You can still inspect 'numeric_summary' and 'summary_pretty' if running interactively.")

