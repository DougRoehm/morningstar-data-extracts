"""
Functions for processing Key Metrics exports from Morningstar.com
"""

import numpy as np
import pandas as pd
import pathlib

### Functions for processing Key Metrics > Financial Summary exports ###

def process_income_statement_from_summary_export(file_path):
    """
    Extract and process the Income Statement from a Morningstar Key Metrics
    Financial Summary export.

    This involves separating the Income Statement data, removing unneeded
    columns, renaming and modifying indexes, transposing the data table,
    changing data types from strings to numbers, modifying % columns from
    whole numbers to decimals and returning a DataFrame. 

    Parameters:
        file_path (obj): Pathlib path to the file including the file name.
        
    Returns:
        DataFrame (obj): A Pandas DataFrame.

    Raises:

    """
    
    # Get company ticker from xls worksheet name
    with pd.ExcelFile(file_path) as xls:
        sheet_names = xls.sheet_names
        company_ticker = sheet_names[0]
    
    # File information
    file_name = str(file_path).split("/")[-1]

    # Convert to Pandas DataFrame
    df_is = pd.read_excel(file_path, index_col=0)

    # Only keep the Income Statement rows and year columns
    df_is = df_is.loc["Revenue":"Dividend Per Share", "2015":"2024"]

    # Rename the index
    df_is.index.rename("Description", inplace=True)

    # Transpose the DataFrame
    df_is_t = df_is.T

    # Get data for new MultiIndex    
    company_ticker = file_name.split("_")[-1].replace(".xls", "")
    is_index = pd.MultiIndex.from_product([[company_ticker], df_is_t.index], 
                                          names=["Ticker", "Year"])
    
    # Create new MultiIndex
    df_is_t.index = is_index

    # Create list of column names for data type conversion
    is_column_names = df_is_t.columns

    # Convert columns to numbers
    for column in is_column_names:
        df_is_t[column] = pd.to_numeric(df_is_t[column], errors='coerce')

    # Convert % columns from whole number to decimal by diving by 100
    df_is_t["Revenue Growth %"] = (
        df_is_t["Revenue Growth %"] / 100).round(4)
    
    df_is_t["Gross Profit Margin %"] = (
        df_is_t["Gross Profit Margin %"] / 100).round(4)
    
    df_is_t["Operating Margin %"] = (
        df_is_t["Operating Margin %"] / 100).round(4)
    
    df_is_t["EBIT Margin %"] = (
        df_is_t["EBIT Margin %"] / 100).round(4)
    
    df_is_t["EBITDA Margin %"] = (
        df_is_t["EBITDA Margin %"] / 100).round(4)
    
    df_is_t["Net Profit Margin %"] = (
        df_is_t["Net Profit Margin %"] / 100).round(4)

    print(f"Income Statement processed from: {file_name}")

    return df_is_t

def process_balance_sheet_from_summary_export(file_path):
    """
    Extract and process the Balance Sheet from a Morningstar Key Metrics
    Financial Summary export.

    This involves separating the Balance Sheet data, removing unneeded
    columns, renaming and modifying indexes, transposing the data table,
    changing data types from strings to numbers, modifying % columns from
    whole numbers to decimals and returning a DataFrame. 

    Parameters:
        file_path (obj): Pathlib path to the file including the file name.
        
    Returns:
        DataFrame (obj): A Pandas DataFrame.

    Raises:

    """
    
    # Get company ticker from xls worksheet name
    with pd.ExcelFile(file_path) as xls:
        sheet_names = xls.sheet_names
        company_ticker = sheet_names[0]
    
    # File information
    file_name = str(file_path).split("/")[-1]
    
    # Convert to Pandas DataFrame
    df_bs = pd.read_excel(file_path, index_col=0)
    
    # Only keep the Income Statement rows and year columns
    df_bs = df_bs.loc["Total Assets": "Debt to Equity", "2015":"2024"]
    
    # Rename the index
    df_bs.index.rename("Description", inplace=True)
    
    # Transpose the DataFrame
    df_bs_t = df_bs.T
    
    # Get data for new MultiIndex    
    company_ticker = file_name.split("_")[-1].replace(".xls", "")
    bs_index = pd.MultiIndex.from_product([[company_ticker], df_bs_t.index], 
                                          names=["Ticker", "Year"])
    
    # Create new MultiIndex
    df_bs_t.index = bs_index
    
    # Create list of column names for data type conversion
    bs_column_names = df_bs_t.columns
    
    # Convert columns to numbers
    for column in bs_column_names:
        df_bs_t[column] = pd.to_numeric(df_bs_t[column], errors='coerce')
    
    # Convert % columns from whole number to decimal by diving by 100
    # no % columns

    print(f"Balance Sheet processed from: {file_name}")

    return df_bs_t

def process_cash_flow_statement_from_summary_export(file_path):
    """
    Extract and process the Cash Flow Statement from a Morningstar Key \n
    Metrics Financial Summary export.

    This involves separating the Cash Flow Statement data, removing \n
    unneeded columns, renaming and modifying indexes, transposing the data \n
    table, changing data types from strings to numbers, modifying % columns \n
    from whole numbers to decimals and returning a DataFrame. 

    Parameters:
        file_path (obj): Pathlib path to the file including the file name.
        
    Returns:
        DataFrame (obj): A Pandas DataFrame.

    Raises:

    """
    
    # Get company ticker from xls worksheet name
    with pd.ExcelFile(file_path) as xls:
        sheet_names = xls.sheet_names
        company_ticker = sheet_names[0]
    
    # File information
    file_name = str(file_path).split("/")[-1]
    
    # Convert to Pandas DataFrame
    df_cf = pd.read_excel(file_path, index_col=0)
    
    # Only keep the Income Statement rows and year columns
    df_cf = df_cf.loc["Cash From Operating Activities":"Change in Cash", 
                      "2015":"2024"]
    
    # Rename the index
    df_cf.index.rename("Description", inplace=True)
    
    # Transpose the DataFrame
    df_cf_t = df_cf.T
    
    # Get data for new MultiIndex    
    company_ticker = file_name.split("_")[-1].replace(".xls", "")
    cf_index = pd.MultiIndex.from_product([[company_ticker], df_cf_t.index], 
                                          names=["Ticker", "Year"])
    
    # Create new MultiIndex
    df_cf_t.index = cf_index
    
    # Create list of column names for data type conversion
    cf_column_names = df_cf_t.columns
    
    # Convert columns to numbers
    for column in cf_column_names:
        df_cf_t[column] = pd.to_numeric(df_cf_t[column], errors='coerce')
    
    # Convert % columns from whole number to decimal by diving by 100
    # no % columns

    print(f"Cash Flow Statement processed from: {file_name}")

    return df_cf_t


### Functions for processing Key Metrics > Profitability and Efficiency exports ###

def process_profitability_and_efficiency_export(file_path):
    """
    Process a Morningstar Key Metrics export of Profitability and
    Efficiency Ratios.

    This involves removing unneeded columns, renaming and modifying
    indexes, modifying % columns from whole numbers to decimals,
    transposing the data table, and saving as a Pandas DataFrame. 

    Parameters:
        file_path (obj): Pathlib path to the file including the file name.
        
    Returns:
        DataFrame (obj): A Pandas DataFrame.

    Raises:

    """
    
    # Get company ticker from xls worksheet name
    with pd.ExcelFile(file_path) as xls:
        sheet_names = xls.sheet_names
        company_ticker = sheet_names[0]
    
    # File information
    file_name = str(file_path).split("/")[-1]
    
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


### Functions for processing Key Metrics > Financial Health exports ###

def process_financial_health_export(file_path):
    """
    Process a Morningstar Key Metrics export of Financial Health Ratios.

    This involves removing unneeded columns, renaming and modifying
    indexes, transposing the data table, changing data types from strings
    to numbers, modifying % columns from whole numbers to decimals and
    saving as a Pandas DataFrame. 

    Parameters:
        file_path (obj): Pathlib path to the file including the file name.
        
    Returns:
        DataFrame (obj): A Pandas DataFrame.

    Raises:

    """
    
    # Get company ticker from xls worksheet name
    with pd.ExcelFile(file_path) as xls:
        sheet_names = xls.sheet_names
        company_ticker = sheet_names[0]
    
    # File information
    file_name = str(file_path).split("/")[-1]
    
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
    # Note any columns that are not year end will still include a month
    df_t.index = df_t.index.set_levels(df_t.index.levels[1].str.replace(
            '-12', '', regex=False),level=1)

    print(f"File processed: {file_name}")

    return df_t


### Functions for processing Key Metrics > Cash Flow Ratios exports ###

def process_cash_flow_ratios_export(file_path):
    """
    Process a Morningstar Key Metrics export of Cash Flow Ratios.

    This involves removing unneeded columns, renaming and modifying
    indexes, transposing the data table, changing data types from strings
    to numbers, modifying % columns from whole numbers to decimals and
    saving as a Pandas DataFrame. 

    Parameters:
        file_path (obj): Pathlib path to the file including the file name.
        
    Returns:
        DataFrame (obj): A Pandas DataFrame.

    Raises:

    """
    
    # Get company ticker from xls worksheet name
    with pd.ExcelFile(file_path) as xls:
        sheet_names = xls.sheet_names
        company_ticker = sheet_names[0]
    
    # File information
    file_name = str(file_path).split("/")[-1]
    
    # Convert to Pandas DataFrame
    df = pd.read_excel(file_path, index_col=0) 

    # Drop unwanted columns
    df.drop(columns=["TTM"], inplace=True)

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
    df_t["Operating Cash Flow Growth % YOY"] = (
        df_t["Operating Cash Flow Growth % YOY"] / 100).round(4)
    
    df_t["Free Cash Flow Growth % YOY"] = (
        df_t["Free Cash Flow Growth % YOY"] / 100).round(4)
    
    df_t["Free Cash Flow/Sales %"] = (
        df_t["Free Cash Flow/Sales %"] / 100).round(4)

    # Remove '-12' from the second level (Year)
    # Note any columns that are not year end will still include a month
    df_t.index = df_t.index.set_levels(df_t.index.levels[1].str.replace(
            '-12', '', regex=False),level=1)

    print(f"File processed: {file_name}")

    return df_t