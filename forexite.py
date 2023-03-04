# %%
# import datetime and dateutil to manipulate dates value
import datetime as dt
from dateutil.relativedelta import relativedelta

# import requests for web scraping
import requests

# import BytesIO and zipfile to get text from zip file
from io import BytesIO
import zipfile

# import pandas and numpy for data manipulation
import pandas as pd
import numpy as np

# %%
# we can get zip file from url in this format: https://www.forexite.com/free_forex_quotes/2021/03/010321.zip

# a function to construct url 
def get_url(date):

    # construct url
    url = f"https://www.forexite.com/free_forex_quotes/{date.year}/{str(date.month).zfill(2)}/{str(date.day).zfill(2)}{str(date.month).zfill(2)}{str(date.year)[-2:]}.zip"
    
    return url

# %%
# a function to construct a dataframe from zip file
def get_df(url):
    
    # get zip file from url
    with requests.get(url) as response:
        
        # use BytesIO to avoid writing zip file to disk
        with zipfile.ZipFile(BytesIO(response.content)) as zip_file:
            
            # get the name of the file in the zip archive 
            filename = zip_file.namelist()[0]
            
            # use pandas to read the file into a dataframe
            with zip_file.open(filename) as text_file:
                df = pd.read_csv(text_file,
                                 sep=',',                                           # seperated by comma
                                 usecols=[0,1,6],                                   # only read 3 columns
                                 names=['Ticker', 'Date', 'Close Price'],           # set the column names
                                 skiprows=1,                                        # skip the first row
                                 parse_dates=[1])                                   # parse the second column as date data type
    return df

# %%
# a function to get inputs from user
def configure():

    # define a list of valid ticker options
    valid_ticker_options = ['a', 's', 'u']
    while True:
        ticker_option = input('Enter "a" to retrieve all currency pairs or "s" to specifiy tickers or "u" to upload a ticker list: ' )
    
        if ticker_option in valid_ticker_options:
            
            if ticker_option == 'a':
                print('Getting all currency pairs... Please note that getting all pairs may result in missing values on non-trading days.')
                
            elif ticker_option == 'u':
                print('Please put your tickers in an excel or csv file with no header and then enter the file path.')
                file_path = input('Enter the file path: ')
                file_path = file_path.replace('\\', '/')
                print(f'Reading tickers from "{file_path}"')
                
            elif ticker_option == 's':
                print('Please specify the tickers.')
                
            break
        
        print('Invalid option. Please enter "a" to retrieve all currency pairs or "s" to specifiy tickers: ')
    
    
    
    if ticker_option == 's':
        tickers = []
        while True:
            ticker = input('Enter ticker (e.g. EURUSD) or "done" to exit: ')
            print(f'{ticker}')
            if ticker == 'done':
                break
            tickers.append(ticker)
    
    elif ticker_option == 'u':
        while True:
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path, header=None)
                break
            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path, header=None)
                break
            elif file_path.endswith('.txt'):
                df = pd.read_csv(file_path, sep=',', header=None)
                break
            
            print('Invalid file type. Please upload an excel or csv file.')
            file_path = input('Enter the file path: ')
            
        tickers = df.iloc[:, 0].tolist()
            
    elif ticker_option == 'a':
        tickers = None
    
    
    
    # define a list of valid frequency options
    valid_frequency_options = ['d', 'm']
    while True:
        frequency = input('Enter "d" for daily data or "m" for monthly data: ')
        
        if frequency in valid_frequency_options:
            if frequency == 'd':
                print('Getting daily data...')
            
            elif frequency == 'm':
                print('Getting monthly data...')

            break
        
        print('Invalid option. Please enter "d" for daily data, "m" for monthly data: ')
    
    
    
    if frequency == 'd':
        while True:
            while True:
                start = input('Enter start date (e.g. 2022-1-1): ')
                if dt.datetime.strptime(start, '%Y-%m-%d') >= dt.datetime.today():
                    print('Start date cannot be today or later than today. Please enter a valid start date.')
                else: 
                    print(f'Start date: {start}')
                    break
        
            while True:
                end = input('Enter end date (e.g. 2022-1-31): ')
                if dt.datetime.strptime(end, '%Y-%m-%d') >= dt.datetime.today():
                    print('End date cannot be today or later than today. Please enter a valid end date.')
                else:
                    print(f'End date: {end}')
                    break
            
            if dt.datetime.strptime(start, '%Y-%m-%d') <= dt.datetime.strptime(end, '%Y-%m-%d'):
                break
           
            print('End date cannot be earlier than start date. Please enter start date and end date again.')
            

    elif frequency == 'm':
        while True:
            
            while True:
                start = input('Enter start year and month (e.g. 2022-1): ')
                if dt.datetime.strptime(start, '%Y-%m').replace(day=1) + relativedelta(months=1, days=-1) >= dt.datetime.today():
                    print('Start month cannot be this month or later than this month. Please enter a valid start date.')
                else:
                    print(f'Start date: {start}')
                    break
            
            while True:
                end = input('Enter end year and month (e.g. 2022-12): ')
                if dt.datetime.strptime(end, '%Y-%m').replace(day=1) + relativedelta(months=1, days=-1) >= dt.datetime.today():
                    print('End month cannot be this month or later than this month. Please enter a valid end date.')
                else:
                    print(f'End date: {end}')
                    break
            
            if dt.datetime.strptime(start, '%Y-%m') <= dt.datetime.strptime(end, '%Y-%m'):
                break
            
            print('End date cannot be earlier than start date. Please enter start date and end date again.')
    
    
    return ticker_option, tickers, frequency, start, end

# %%
def forex_monthly(ticker_option, tickers, start, end):
    
    ticker_option = ticker_option
    tickers = tickers
    
    # transform input values into date object
    start_date = dt.datetime.strptime(start, '%Y-%m').replace(day=1).date()
    end_date = dt.datetime.strptime(end, '%Y-%m').replace(day=1).date()
    
    # create an empty list to store dataframe
    df_list = []
    
    # loop over each month between start_date and end_date
    while start_date <= end_date:
        
        # calculate the last day of the current month by adding a month and subtracting a day
        last_day = start_date + relativedelta(months=1, days=-1)            
        
        # print the message
        print(f'Processing: {last_day}')
        
        # get url of last day
        url = get_url(last_day)
        
        # get dataframe of the last day
        df = get_df(url)
        
        # filter the df with date, and then get the last row of each ticker
        monthly_close_all = df[df['Date']==np.datetime64(last_day)].groupby(['Ticker']).last().reset_index()
        
        # if user chose to retrieve all currency pairs
        if ticker_option == 'a':
            # use monthly_close_all as final output without handling missing values on non-trading days 
            df_list.append(monthly_close_all)
        
        # if user specified currency pairs or upload a list of tickers
        else:
            # loop over each ticker
            for ticker in tickers:
                
                # filter monthly_close_all with ticker
                ticker_monthly_close = monthly_close_all[monthly_close_all['Ticker']==ticker]

                # before handling missing values, create counters to end while loops
                previous_tries = 0
                following_tries = 0
                date = last_day
                
                # if df is empty, then try retrieving the most recent 10 days data (the previous 5 days and the following 5 days)
                while ticker_monthly_close.empty:
                    # print the message
                    print(f'{ticker} on {date} is empty')
                    
                    # update counter
                    previous_tries += 1
            
                    # start to try the previous 5 days
                    date = last_day + relativedelta(days=-previous_tries)
                    
                    # print the message
                    print(f'Processing: {ticker}, {date}')
            
                    # get url and dataframe
                    url = get_url(date)
                    df = get_df(url)
            
                    # filter and get the last row
                    ticker_monthly_close = df[(df['Ticker']==ticker) & (df['Date']==np.datetime64(date))].tail(1)
                    
                    # if the previous 5 days have no data, then try the following 5 days
                    if previous_tries == 5:
                        while ticker_monthly_close.empty:
                            print(f'{ticker} on {date} is empty')
                            
                            # update counter
                            following_tries += 1
                            
                            # update the date
                            date = last_day + relativedelta(days=following_tries)
                        
                            # print the message
                            print(f'Processing: {ticker}, {date}')
                        
                            # get url and dataframe
                            url = get_url(date)
                            df = get_df(url)
            
                            # filter and get the last row
                            ticker_monthly_close = df[(df['Ticker']==ticker) & (df['Date']==np.datetime64(last_day))].tail(1)
                            
                            # if the following 5 days have no data as well, then possibly this ticker contains a cryptocurrency that hadn't been invented yet
                            if following_tries == 5:
                                break
                            
                    # break the loop
                    if following_tries == 5:
                            break                             
                
                # add null values to empyty currency pair
                if ticker_monthly_close.empty:
                    ticker_monthly_close = pd.concat([ticker_monthly_close, pd.DataFrame([[ticker, np.datetime64(last_day), np.nan]], columns=ticker_monthly_close.columns)])
                
                # append the dataframe of each month to the list
                df_list.append(ticker_monthly_close)      
    
        # move to the next month
        start_date = start_date + relativedelta(months=1)
        
    # concate all dataframes in the list
    df = pd.concat(df_list, ignore_index=True)
    
    # sort the dataframe by tickers and dates
    df = df.sort_values(by=['Ticker', 'Date'])

    # write df to an excel file
    excel_file_name = f'{start}_{end} monthly {dt.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    df.to_excel(excel_file_name, sheet_name='monthly', index=False)
    
    # print the message
    print(f'Excel file {excel_file_name} generated')

# %%
def forex_daily(ticker_option, tickers, start, end):
    
    ticker_option = ticker_option
    tickers = tickers
        
    # define start_date
    start_date = dt.datetime.strptime(start, "%Y-%m-%d").date()
    
    # subtracting 5 days from the start date for interpolation purpose
    start_date_for_itp = start_date + relativedelta(days=-5)
    
    # define end_date 
    end_date = dt.datetime.strptime(end, "%Y-%m-%d").date()
    
    # add 5 days to the end date for interpolation purpose
    end_date_for_itp = end_date + relativedelta(days=5)
    
    # create an empty list to store dataframe
    df_list = []
    
    # loop over each day between start_date and end_date
    while start_date_for_itp <= end_date_for_itp:
        
        # print the message
        print(f'Processing: {start_date_for_itp}')
        
        # get url
        url = get_url(start_date_for_itp)
        
        # get dataframe
        df = get_df(url)
        
        # filter the df with date, and then get the last row of each ticker
        daily_close_all = df[df['Date']==np.datetime64(start_date_for_itp)].groupby(['Ticker']).last().reset_index()
        
        # if user chose to retrieve all currency pairs
        if ticker_option == 'a':
            # use daily_close_all as final output without handling missing values on non-trading days
            df_list.append(daily_close_all)
        
        # if user specified tickers or upload a list of tickers
        else:
            # loop over each ticker
            for ticker in tickers:
                
                # filter daily_close_all with ticker
                ticker_daily_close = daily_close_all[(daily_close_all['Ticker']==ticker)]
                
                # if there is no data on this date, put null value on this date
                if ticker_daily_close.empty:
                    ticker_daily_close = pd.concat([ticker_daily_close, pd.DataFrame([[ticker, np.datetime64(start_date_for_itp), np.nan]], columns=ticker_daily_close.columns)])

                # append the dataframe of each day to the list
                df_list.append(ticker_daily_close)      
    
        # move to the next day
        start_date_for_itp = start_date_for_itp + relativedelta(days=1)
        
    # concating all dataframes in the list
    df = pd.concat(df_list, ignore_index=True)
    
    # if user specified currency pairs, then we can interpolate null values
    if ticker_option == 's' or ticker_option == 'u':
        
        # interpolate the NaN values using linear interpolation
        df['Interpolated Price'] = df.groupby(['Ticker'], group_keys=False)['Close Price'].apply(lambda x: x.interpolate(method='linear')).round(4)
    
    # filter out the days used for interpolation
    df = df.loc[(df['Date']>=np.datetime64(start_date)) & (df['Date']<=np.datetime64(end_date))]
    
    # sort the df by tickers and dates
    df = df.sort_values(by=['Ticker', 'Date'])

    # write df to an excel file
    excel_file_name = f'{start}_{end} daily {dt.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    df.to_excel(excel_file_name, sheet_name='daily', index=False)
    
    # print the message
    print(f'Excel file {excel_file_name} generated')

# %%
def forexite():
    # get user inputs from the configure() function
    ticker_option, tickers, frequency, start, end = configure()
    
    # calling functions
    if frequency == 'd':
        if ticker_option == 'a':
            forex_daily(ticker_option, None, start, end)

        elif ticker_option == 'u':
            forex_daily(ticker_option, tickers, start, end)
            
        else:
            tickers = list(tickers)
            forex_daily(ticker_option, tickers, start, end)
    
    else:
        if ticker_option == 'a':
            forex_monthly(ticker_option, None, start, end)

        elif ticker_option == 'u':
            forex_monthly(ticker_option, tickers, start, end)

        else:
            tickers = list(tickers)
            forex_monthly(ticker_option, tickers, start, end)

            
# %%
# run the function
forexite()

