# morningstar-data-extracts
This repository is used for processing free financial data extracted from Morningstar.com

Morningstar provides various financial data on public companies. For example if you are wanting information on Apple Inc. you would enter the companies name or ticker (AAPL) in the search box to be directed to the company's financial information on Morningstar. Different types of financial information is provided on different tabs/pages for each company. The various different pages of financial information can then be exported and saved as separate excel (xls) files.

Often times the information in the Key Metrics section is used for comparing different publicly traded companies to each other. 

The programs in this repository can be used to convert the xls file exports from the Morningstar website. The following exports currently can be processed:
- Key Metrics > Financial Summary
- Key Metrics > Profitability and Efficiency
- Key Metrics > Financial Health

The program expects a path to a folder containing the files to be processed. Each type of export will need to be in a separate folder. The program will then take all the files in the folder, convert them to Pandas DataFrames, make various modification, combine them into one DataFrame and save the file as a csv file in the output folder.

See the demonstration folder for an example of various telecommunications guideline companies being combined for each of the above export files.

**Note: When exporting the data from Morningstar each file has manually had an underscore and ticker added.** For example using the company above "_AAPL" would be added to the end of each morningstar export file. This helps for file management and **is currently expected in the way the code is written.** This could be changed at some point however, because when Morningstar exports the xls file it names the tab using the company ticker. The worksheet tab could be used to extract the ticker instead of expecting it be at the end of each file.