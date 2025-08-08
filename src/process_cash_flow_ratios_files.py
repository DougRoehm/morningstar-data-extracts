"""
Used to process a folder containing Morningstar Cash Flow Ratios exports
"""

import os
import sys
import pathlib
import numpy as np
import pandas as pd
from general_functions import get_files_from_folder
from cash_flow_ratios_functions import process_cash_flow_ratios_export


# Lists to be populated
industry_description = ""
gpc_tickers = []
file_names = []

# Collect user information
industry_description = input("Enter industry description:\n").lower().replace(" ", "-")
folder_path = input("Enter folder path:\n")
file_paths = get_files_from_folder(folder_path)

# Root, source, and output folders
current_file = pathlib.Path(__file__).resolve()
project_root = current_file.parent.parent
src_folder = project_root / 'src'
output_folder = project_root / "output"

# Process files using financial health functions and save to DataFrame
df_ratios = pd.DataFrame()

for file in file_paths:
    df_temp = process_cash_flow_ratios_export(file)
    df_ratios = pd.concat([df_ratios, df_temp], axis = 0)

# Output file name and path
output_file_name = f"gpc-cash-flow-ratios-{industry_description}.csv"
output_file_path = output_folder / output_file_name

# Save file
df_ratios.to_csv(output_file_path)