My code consists of 2 scripts:
Retrieving_data.py
Creating_Output_Spreadsheet.py

NB: THIS IS AN EXERCISE IN DATA MANIPULATION AND API USE, HENCE THE MISSING DATA ETC.

------------------

Instructions:

1. 1. Create a folder to hold the equity data we will retrieve from the
API (in the form of xlsx files). It is currently called 'Data' in my code.

2. Run Retrieving_data.py. Ensure to update the file paths/names on lines 12 and 89. 

3. Run Creating_Output_Spreadsheet.py. Ensure to update the file paths/names on lines 15, 23, 29, 44 and 52.


Retrieving_data.py pulls data from Yahoo Finance using the yfinance Python package. This script downloads the data and saves the previous four 
earning dates to an Excel workbook for each stock.

Creating_Output_Spreadsheet.py reads the individual Excel sheets containing the earning dates for each stock and writes it to a single Excel workbook.

N/D = No Data. This indicates the data is unavailable because Yahoo Finance does not contain it. I apologise for the significant missing data, 
Yahoo Finance is the only free API I could find that contains quarterly data. 

