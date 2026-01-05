# this file is to used to combine multiple period reports into one file
# Function:
# Will combine our master logs from each period into one master log
# Will combine our weekly summaries from each period into one weekly summary
# Will create a new excel file with the combined master log and weekly summary

import pandas as pd
from skyepipeline_files.MasterLogCreation import build_master_log
from skyepipeline_files.WeeklySummaryCreator import build_weekly_summary
from skyepipeline_files.BuildWeeklyWorkbook import build_weekly_workbook
import os
from datetime import datetime

def combine_period_reports(report_files, output_path):
    combined_master_df = pd.DataFrame()
    combined_summary_df = pd.DataFrame()

    for report_file in report_files:
        # Read the master log and weekly summary from each report file
        with pd.ExcelFile(report_file) as xls:
            master_df = pd.read_excel(xls, sheet_name='Master Log')
            summary_df = pd.read_excel(xls, sheet_name='Weekly Summary')

            # Append to the combined DataFrames
            combined_master_df = pd.concat([combined_master_df, master_df], ignore_index=True)
            combined_summary_df = pd.concat([combined_summary_df, summary_df], ignore_index=True)

    # Write the combined DataFrames to a new Excel file
    with pd.ExcelWriter(output_path) as writer:
        combined_master_df.to_excel(writer, sheet_name='Combined Master Log', index=False)
        combined_summary_df.to_excel(writer, sheet_name='Combined Weekly Summary', index=False)

    print(f"\n✅ Done! Combined report written to: {os.path.abspath(output_path)}")
    
if __name__ == "__main__":
    # reads in a folder where we can input our path into a variable below (folder will consist only of the period reports)

    report_folder = "/path/to/period/reports"
    output_file = "/path/to/output/Combined_Skye_Period_Report.xlsx"
    report_files = [os.path.join(report_folder, f) for f in os.listdir(report_folder) if f.endswith('.xlsx')]
    combine_period_reports(report_files, output_file)
  