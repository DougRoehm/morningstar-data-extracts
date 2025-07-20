# morningstar-data-extracts
This repository is used for processing free financial data extracted from Morningstar.com

Morningstar provides various financial data on public companies. For example if you are wanting information on Apple Inc. you would enter the companies name or ticker (AAPL) in the search box to be directed to the company's financial information on Morningstar. Different types of financial information is provided on different tabs/pages for each company. The various different pages of financial information can then be exported and saved as separate excel (xls) files.

Often this information is used for comparing many different publicly traded companies.

The programs in this repository can be used to convert the various xls files exported by Morningstar into Pandas DataFrames. These DataFrames can then be easily combined and used for financial analysis projects.