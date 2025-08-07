"""
Used to process a folder containing Morningstar Financial Summary exports
Found on Morningstar.com under Key Metrics > Financial Summary
"""

import os
import sys
import pathlib
import numpy as np
import pandas as pd
from general_functions import get_files_from_folder
from financial_summary_functions import (
    process_income_statement_from_summary_export,
    process_balance_sheet_from_summary_export,
    process_cash_flow_statement_from_summary_export
)

# Collect user information
industry_description = input("Enter industry description:\n").lower().replace(" ", "-")
folder_path = input("Enter folder path:\n")
file_paths = get_files_from_folder(folder_path)

# Root, source, and output folders
current_file = pathlib.Path(__file__).resolve()
project_root = current_file.parent.parent
src_folder = project_root / 'src'
output_folder = project_root / "output"


## Process Financial Statements ##

# Process Income Statement using function and save to DataFrame
df_income_statement = pd.DataFrame()

for file in file_paths:
    df_temp = process_income_statement_from_summary_export(file)
    df_income_statement = pd.concat([df_income_statement, df_temp], axis = 0)

# Process Balance Sheet using function and save to DataFrame
df_balance_sheet = pd.DataFrame()

for file in file_paths:
    df_temp = process_balance_sheet_from_summary_export(file)
    df_balance_sheet = pd.concat([df_balance_sheet, df_temp], axis = 0)

# Process Cash Flow Statement using function and save to DataFrame
df_cash_flow_statement = pd.DataFrame()

for file in file_paths:
    df_temp = process_cash_flow_statement_from_summary_export(file)
    df_cash_flow_statement = pd.concat([df_cash_flow_statement, df_temp], axis = 0)


## Save Financial Statements ##

# Output file names and paths
output_file_name_is = f"gpc-summary-income-statement-{industry_description}.csv"
output_file_path_is = output_folder / output_file_name_is

output_file_name_bs = f"gpc-summary-balance-sheet-{industry_description}.csv"
output_file_path_bs = output_folder / output_file_name_bs

output_file_name_cf = f"gpc-summary-cash-flow-{industry_description}.csv"
output_file_path_cf = output_folder / output_file_name_cf

# Save files
df_income_statement.to_csv(output_file_path_is)
df_balance_sheet.to_csv(output_file_path_bs)
df_cash_flow_statement.to_csv(output_file_path_cf)