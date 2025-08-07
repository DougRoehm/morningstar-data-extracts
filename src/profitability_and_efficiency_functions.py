"""
Function for processing Morningstar export of: 
Key Metrics > Profitability and Efficiency Ratios
"""

import numpy as np
import pandas as pd

def process_profitability_and_efficiency_export(file_path):
    """
    Process a Morningstar Key Metrics export of Profitability and
    Efficiency Ratios.

    This involves removing unneeded columns, renaming and modifying
    indexes, modifying % columns from whole numbers to decimals,
    transposing the data table, and saving as a Pandas DataFrame. 

    Parameters:
        file_path (str): The path to the file including the file name.
        
    Returns:
        DataFrame (obj): A Pandas DataFrame.

    Raises:

    """
    
    # File information
    file_name = str(file_path).split("/")[-1]
    company_ticker = file_name.split("_")[-1].replace(".xls", "")

    # Convert to Pandas DataFrame
    df = pd.read_excel(file_path, index_col=0) 

    # Drop last two columns
    df.drop(columns=["Current", "5-Yr"], inplace=True)

    # Rename Index to Ratio
    df.index.rename("Ratio", inplace=True)

    # Transpose DataFrame
    df_t = df.T

    # Create new MultiIndex
    new_index = pd.MultiIndex.from_product([[company_ticker], df_t.index], 
                                           names=["Ticker", "Year"])

    # Assign the new index to the DataFrame
    df_t.index = new_index

    # Create list of column names for data type conversion
    column_names = df_t.columns

    # Convert columns to numbers
    for column in column_names:
        df_t[column] = pd.to_numeric(df_t[column], errors='coerce')
    
    # Convert % columns from whole number to decimal
    df_t["Return on Asset %"] = (
        df_t["Return on Asset %"] / 100).round(4)
    
    df_t["Return on Equity %"] = (
        df_t["Return on Equity %"] / 100).round(4)
    
    df_t["Return on Invested Capital %"] = (
        df_t["Return on Invested Capital %"] / 100).round(4)
    
    print(f"File processed: {file_name}")

    return df_t