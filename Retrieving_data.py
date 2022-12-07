import pandas as pd
import numpy as np
import requests as rq
import json
from os.path import exists
import yfinance as yf

import datetime
from datetime import datetime as dt

#Load in the universe spreadsheet as a dataframe
universe_df = pd.read_excel(r'/Users/kelvinbrinham/Desktop/Python_practice/Aperture_Task_1/Universe.xlsx')

#List of Tickers from spreadsheet
Ticker_list = list(universe_df['Ticker'])
#List of Tickers minus exchanges
Ticker_list_stripped = []

#Strip the ticker symbols to the base tickers
for Ticker in Ticker_list:
    Ticker_list_stripped.append(Ticker.split()[0])

#------------------------
#Get data from Yahoo Finance
tickers_data = yf.Tickers(Ticker_list_stripped)
set_of_invalid_Tickers = set()
#Yahoo has no data on these stocks and throws a key error as a result
set_of_invalid_Tickers.update(['HEIA', 'VIE', 'PUB', 'ADS', 'DPW', 'VER'])



#Retrieve earning dates
if __name__ == '__main__':
    for Ticker in Ticker_list_stripped:
        # print(Ticker)
        #Check if ticker is invalid (i.e. Yahoo doesnt accept it)
        if Ticker in set_of_invalid_Tickers:
            continue

        #Check if ticker is empty
        if not Ticker:
            continue

        #Check if Ticker data exists
        Stock_data = tickers_data.tickers[Ticker] #Data for individual stock
        if Stock_data is None:
            # print('No Stock Data')
            continue

        Stock_Earning_Dates_Data = Stock_data.earnings_dates #Earning dates for individual stock
        if Stock_Earning_Dates_Data is None:
            # print('No Stock Earning Date Data')
            continue

        #Add index column to dataframe
        Stock_Earning_Dates_Data = Stock_Earning_Dates_Data.reset_index()
        #Retrieve Earning dates
        Stock_Earning_Dates = Stock_Earning_Dates_Data['Earnings Date']

        Stock_Earning_Dates_Valid = []
        #Remove non 2022 earning dates, this could be changed to include historical data
        for i in range(len(Stock_Earning_Dates)):
            #Retrieve just the year from the datetime
            year = str(Stock_Earning_Dates_Data['Earnings Date'][i].to_period('Y'))
            if year == '2022':
                Earning_date = Stock_Earning_Dates_Data['Earnings Date'][i].to_pydatetime().astimezone().replace(tzinfo=None)
                Stock_Earning_Dates_Valid.append(Earning_date)

        #Check that Yahoo has 2022 data for this stock
        if not Stock_Earning_Dates_Valid:
            # print('No 2022 Earning Date Data')
            continue

        #Ensure Yahoo has 4 2022 earning dates
        #NB: Some of the stocks contain more than 4 2022 dates
        if len(Stock_Earning_Dates_Valid) != 4:
            #Insufficient data
            continue


        Stock_Earning_Dates_Valid.reverse()
        # Stock_Earning_Dates_Valid.sort()
        #Create a data frame for an individual stock containing the earning dates
        data = {'Financial Report Release Date': None}
        Stock_Earning_Dates_DataFrame = pd.DataFrame(data, index = ['2021 Q4', '2022 Q1', '2022 Q2', '2022 Q3'])
        Stock_Earning_Dates_DataFrame['Financial Report Release Date'] = Stock_Earning_Dates_Valid

        #Write the earning dates for an individual stocks to an Excel sheet
        Stock_Earning_Dates_DataFrame.to_excel(f'/Users/kelvinbrinham/Desktop/Python_practice/Aperture_task_1_Yahoo/Output_quarterly/{Ticker}.xlsx', sheet_name='Sheet1')
