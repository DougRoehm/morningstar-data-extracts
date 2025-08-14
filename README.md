# morningstar-data-extracts
This repository is used for processing free Key Metric financial data 
extracted from [Morningstar.com](https://morningstar.com)


## About:

This repo is used to process and combine Morningstar Key Metric information for 
multiple companies, that is then used for conducting Financial Statement Ratio 
Analysis. Ratio analysis can be used to help quantify some of the strengths and 
weaknesses when comparing one company to another or to a subset of companies. For 
example, profitability ratios such as profit margin (profit / sales) and return on 
assets (net income / total assets) can be used to compare the differing levels of 
profitability between companies.

These ratios can be computed manually from a company's financial statements, however 
financial services providers like Morningstar compute many of these Financial 
Statement Ratios for publicly traded companies and provide them as part of their 
service packages. Morningstar provides many of these ratios for free, however they 
must be exported company by company and there are several different export files for 
each company. This is where this repo comes into play as its designed to process 
these exports and combine the many company exports into one file for each export type.


## How to get to extract data on Morningstar.com:

Using Apple Inc as an example. To extract key metric financial data navigate to 
Morningstar.com, enter Apple Inc or AAPL (the company's ticker) into the search bar to 
bring up the company's page on Morningstar, then click on the Key Metrics tab. From 
there different types of Key Metrics data can be exported as excel (xls) files. 


## Types of Key Metrics exports that can be processed:

This repository can be used to process and combine the following types of Key Metric 
exports:
- Key Metrics > Financial Summary
- Key Metrics > Profitability and Efficiency
- Key Metrics > Financial Health
- Key Metrics > Cash Flow Ratios


## Walkthrough of using the repository:

The repository will ask the user for the following inputs:
1. What type of Key Metric export they want to process.
2. An industry description to add to the end of the processed file.
3. To copy and paste the path to the folder containing files to be processed.

Note: each type of export will need to be in a separate folder. The program will then take all the files in the folder, convert them to Pandas DataFrames, make various 
modifications, combine them into one DataFrame and save the file as a csv file in the 
output folder.

See the files in the demonstration folder for an example of various telecommunications 
guideline companies being combined for each fo the above export file types.