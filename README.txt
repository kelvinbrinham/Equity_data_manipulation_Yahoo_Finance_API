My code consists of 2 scripts:
Retrieving_data.py
Creating_Output_Spreadsheet.py

NB: THIS IS AN EXERCISE IN DATA MANIPULATION AND API USE (USING A FREE API), HENCE THE MISSING DATA ETC.

------------------

Instructions:

1. Run Retrieving_data.py. The file paths should work within this working directory. I've included the data this code produces to save time
if you want to try out the second script on it's own. Just make sure to change the file paths in Creating_Output_Spreadsheet.py. to Data_Example rather than Data.

2. Run Creating_Output_Spreadsheet.py. 


Retrieving_data.py pulls data from Yahoo Finance using the yfinance Python package. This script downloads the data and saves the previous four 
earning dates to an Excel workbook for each stock.

Creating_Output_Spreadsheet.py reads the individual Excel sheets containing the earning dates for each stock and writes it to a single Excel workbook.

N/D = No Data. This indicates the data is unavailable because Yahoo Finance does not contain it. I apologise for the significant missing data, 
Yahoo Finance is the only free API I could find that contains quarterly data. 

