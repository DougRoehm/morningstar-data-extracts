"""
Functions for processing Morningstar export of: 
Key Metrics > Financial Health Ratios
"""

import numpy as np
import pandas as pd

def process_financial_health_export(file_path):
    """
    Process a Morningstar Key Metrics export of Financial Health Ratios.

    This involves removing unneeded columns, renaming and modifying
    indexes, transposing the data table, changing data types from strings
    to numbers, modifying % columns from whole numbers to decimals and
    saving as a Pandas DataFrame. 

    Parameters:
        file_path (str): The path to the file including the file name.
        
    Returns:
        DataFrame (obj): A Pandas DataFrame.

    Raises:

    """
    
    # File information
    # file_name = file_path.split("/")[-1]
    # company_ticker = file_name.split("_")[-1].replace(".xls", "")

    # File information from PosixPath object
    # TODO change this to get the company ticker using the
    # TODO Pandas object created below from the tab name
    file_name = str(file_path).split("/")[-1]
    company_ticker = file_name.split("_")[-1].replace(".xls", "")

    # Convert to Pandas DataFrame
    df = pd.read_excel(file_path, index_col=0) 

    # Drop unwanted columns
    df.drop(columns=["Latest Qtr"], inplace=True)

    # Rename Index to Ratio
    df.index.rename("Ratio", inplace=True)

    # Transpose DataFrame
    df_t = df.T

    # Create new MultiIndex
    new_index = pd.MultiIndex.from_product(
        [[company_ticker], df_t.index], names=["Ticker", "Year"])

    # Assign the new index to the DataFrame
    df_t.index = new_index

    # Create list of column names for data type conversion
    column_names = df_t.columns

    # Convert columns to numbers
    for column in column_names:
        df_t[column] = pd.to_numeric(df_t[column], errors='coerce')
    
    # Convert % columns from whole number to decimal
    df_t["Cap Ex as a % of Sales"] = (
        df_t["Cap Ex as a % of Sales"] / 100).round(4)

    # Remove '-12' from the second level (Year)
    df_t.index = df_t.index.set_levels(df_t.index.levels[1].str.replace(
            '-12', '', regex=False),level=1)

    print(f"File processed: {file_name}")

    return df_t