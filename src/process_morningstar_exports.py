"""
This is the main program to process Morningstar Key Metrics exports.
"""

import os
import sys
import pathlib
import numpy as np
import pandas as pd
# from general_functions import get_files_from_folder
from morningstar_functions import (
    process_income_statement_from_summary_export,
    process_balance_sheet_from_summary_export,
    process_cash_flow_statement_from_summary_export,
    process_profitability_and_efficiency_export,
    process_financial_health_export,
    process_cash_flow_ratios_export
)

def main():

    # Collect user information
    print("This program is meant to process exports from Morningstar.com")
    print("The following types of exports can be processed:")
    print("  1. Key Metrics > Financial Summary")
    print("  2. Key Metrics > Profitability and Efficiency")
    print("  3. Key Metrics > Financial Health")
    print("  4. Key Metrics > Cash Flow Ratios")
    print("")
    export_type = get_user_export_type()
    industry_description = input(
        "Enter an industry description to add to the end of the processed file:\n"
            ).strip().lower().replace(" ", "-")
    folder_path = input(
        "Copy and paste the path to the folder containing files to be processed:\n")
    file_paths = get_files_from_folder(folder_path)

    # Root, source, and output folders
    current_file = pathlib.Path(__file__).resolve()
    project_root = current_file.parent.parent
    src_folder = project_root / 'src'
    output_folder = project_root / "output"

    # Process files and save output
    if export_type == 1:
        print("Processing Key Metric > Financial Summary files...")
        # Process Income Statement using function and save to DataFrame
        df_income_statement = pd.DataFrame()

        for file in file_paths:
            df_temp = process_income_statement_from_summary_export(file)
            df_income_statement = pd.concat([df_income_statement, df_temp], axis = 0)

        # Save DataFrame as csv file in output folder
        output_file_name_is = f"gpc-summary-income-statement-{industry_description}.csv"
        output_file_path_is = output_folder / output_file_name_is
        df_income_statement.sort_index().to_csv(output_file_path_is)

        # Process Balance Sheet using function and save to DataFrame
        df_balance_sheet = pd.DataFrame()

        for file in file_paths:
            df_temp = process_balance_sheet_from_summary_export(file)
            df_balance_sheet = pd.concat([df_balance_sheet, df_temp], axis = 0)

        # Save DataFrame as csv file in output folder
        output_file_name_bs = f"gpc-summary-balance-sheet-{industry_description}.csv"
        output_file_path_bs = output_folder / output_file_name_bs
        df_balance_sheet.sort_index().to_csv(output_file_path_bs)

        # Process Cash Flow Statement using function and save to DataFrame
        df_cash_flow_statement = pd.DataFrame()

        for file in file_paths:
            df_temp = process_cash_flow_statement_from_summary_export(file)
            df_cash_flow_statement = pd.concat([df_cash_flow_statement, df_temp], axis = 0)
    
        # Save DataFrame as csv file in output folder
        output_file_name_cf = f"gpc-summary-cash-flow-{industry_description}.csv"
        output_file_path_cf = output_folder / output_file_name_cf
        df_cash_flow_statement.sort_index().to_csv(output_file_path_cf)

    else:
        pass

    if export_type == 2:
        print("Processing Key Metric > Profitability and Efficiency files...")
        # Process files using profitability and efficiency functions
        df_ratios = pd.DataFrame()

        for file in file_paths:
            df_temp = process_profitability_and_efficiency_export(file)
            df_ratios = pd.concat([df_ratios, df_temp], axis = 0)

        # Save DataFrame as csv file in output folder
        output_file_name = f"gpc-profitability-and-efficiency-{industry_description}.csv"
        output_file_path = output_folder / output_file_name
        df_ratios.sort_index().to_csv(output_file_path)

    else:
        pass

    if export_type == 3:
        # Process files using financial health functions and save to DataFrame
        print("Processing Key Metric > Financial Health files...")
        df_ratios = pd.DataFrame()

        for file in file_paths:
            df_temp = process_financial_health_export(file)
            df_ratios = pd.concat([df_ratios, df_temp], axis = 0)
        
        # Save DataFrame as csv file in output folder
        output_file_name = f"gpc-financial-health-{industry_description}.csv"
        output_file_path = output_folder / output_file_name
        df_ratios.sort_index().to_csv(output_file_path)

    else:
        pass

    if export_type == 4:
        print("Processing Key Metric > Cash Flow Ratios files...")
        # Process files using cash flow ratios functions
        df_ratios = pd.DataFrame()

        for file in file_paths:
            df_temp = process_cash_flow_ratios_export(file)
            df_ratios = pd.concat([df_ratios, df_temp], axis = 0)

        # Save DataFrame as csv file in output folder
        output_file_name = f"gpc-cash-flow-ratios-{industry_description}.csv"
        output_file_path = output_folder / output_file_name
        df_ratios.sort_index().to_csv(output_file_path)

    else:
        pass


def get_user_export_type():
    """Gets export type from user represented by integer between 1 and 4."""
    while True:
        try:
            selection = int(
                input("Enter the number of the type of export you want to process:\n"))
            if 1 <= selection <= 4:
                return selection
            else:
                raise ValueError
        except ValueError:
            print("You must enter a number between 1 and 4.\n")


def get_files_from_folder(folder_path):
    """Gets list of files in provided folder path"""
    folder = pathlib.Path(folder_path)
    return list(folder.glob("*.xls"))


if __name__ == "__main__":
    main()
 