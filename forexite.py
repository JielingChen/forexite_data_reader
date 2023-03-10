
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
# define a function to construct the url for a given date
def get_url(date):

    # construct url
    url = f"https://www.forexite.com/free_forex_quotes/{date.year}/{str(date.month).zfill(2)}/{str(date.day).zfill(2)}{str(date.month).zfill(2)}{str(date.year)[-2:]}.zip"
    
    return url

# %%
# define a function to construct a dataframe from zip file
def get_df(url):
    
    # get zip file from url
    with requests.get(url) as response:
        
        # use BytesIO to avoid writing zip file to computer disk
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
# define a function to get inputs from user
def configure():

    # get ticker option from user
    valid_ticker_options = ['all', 'type', 'upload']
    while True:
        ticker_option = input('Enter "all" to retrieve all currency pairs or "type" to type in tickers or "upload" to upload a list of tickers: ' )
        print(f'Ticker option: {ticker_option}')
        
        if ticker_option in valid_ticker_options:
            
            if ticker_option == 'all':
                print('Getting all currency pairs... Please note that getting all pairs may result in missing values on non-trading days.')
                
            elif ticker_option == 'type':
                print('Please type in the tickers.')
            
            elif ticker_option == 'upload':
                print('Please put your tickers in an excel or csv file with no header and then enter the file path.')
                # check if file type is valid
                valid_extensions = ['.xlsx', '.csv', '.txt']
                while True:
                    file_path = input('Enter the file path: ')
                    print(f'Reading tickers from {file_path}')
                    file_path = file_path.replace('\\', '/')
                    if file_path.endswith(tuple(valid_extensions)):
                        break
                    else:
                        print('Invalid file type. Please upload an excel or csv file.')
                        
            break
        
        print('Enter "all" to retrieve all currency pairs or "type" to type in tickers or "upload" to upload a list of tickers: ')
    
    
    # get tickers from user
    if ticker_option == 'type':
        tickers = []
        while True:
            ticker = input('Enter ticker (e.g. EURUSD) or "done" to exit: ')
            print(f'{ticker}')
            if ticker == 'done':
                break
            tickers.append(ticker)
    
    
    elif ticker_option == 'upload':
        # check if file is valid
        while True:
            try:
                if file_path.endswith('.xlsx'):
                    df = pd.read_excel(file_path, header=None)
                elif file_path.endswith('.csv'):
                    df = pd.read_csv(file_path, header=None)
                elif file_path.endswith('.txt'):
                    df = pd.read_csv(file_path, sep=',', header=None)
                break
            
            except:
                print('Error reading file. Please check the file.')
                file_path = input('Enter the file path: ')
        
        # get ticker list from file  
        tickers = df.iloc[:, 0].tolist()
        print(f'Tickers uploaded: ')
        print(tickers)
    
            
    elif ticker_option == 'all':
        tickers = None
    
    
    # get frequency option from user
    valid_frequency_options = ['d', 'm']
    while True:
        frequency = input('Enter "d" for daily data or "m" for month-end data: ')
        
        if frequency in valid_frequency_options:
            if frequency == 'd':
                print('Getting daily data...')
            
            elif frequency == 'm':
                print('Getting month-end data...')

            break
        
        print('Invalid option. Please enter "d" for daily data, "m" for month-end data: ')
    
    
    # get start and end date from user
    if frequency == 'd':
        while True:
            start = input('Enter start date (e.g. 2022-1-1): ')
            print(f'Start date: {start}')
            try:
                start_date = dt.datetime.strptime(start, '%Y-%m-%d').date()
                if start_date >= dt.datetime.today():
                    print('Start date cannot be today or later than today. Please enter a valid start date.')
                else: 
                    break
            except:
                print('Invalid date format. Please enter a valid start date.')
        
        while True:
            end = input('Enter end date (e.g. 2022-1-31): ')
            print(f'End date: {end}')
            try:
                end_date = dt.datetime.strptime(end, '%Y-%m-%d').date()
                if end_date >= dt.datetime.today():
                    print('End date cannot be today or later than today. Please enter a valid end date.')
                elif end_date < start_date:
                    print('End date cannot be earlier than start date. Please enter a valid end date.')
                else:
                    break
            except:
                print('Invalid date format. Please enter a valid end date.')
            
            
    elif frequency == 'm':
        while True:
            start = input('Enter start year and month (e.g. 2022-1): ')
            print(f'Start date: {start}')
            try:
                # get the last day of the month
                start_date = dt.datetime.strptime(start, '%Y-%m').replace(day=1) + relativedelta(months=1, days=-1)
                if start_date >= dt.datetime.today():
                    print('Start month cannot be this month or later than this month. Please enter a valid start date.')
                else:
                    break
            except:
                print('Invalid date format. Please enter a valid start date.')
            
        while True:
            end = input('Enter end year and month (e.g. 2022-12): ')
            print(f'End date: {end}')
            try:
                # get the last day of the month
                end_date = dt.datetime.strptime(end, '%Y-%m').replace(day=1) + relativedelta(months=1, days=-1)
                if end_date >= dt.datetime.today():
                    print('End month cannot be this month or later than this month. Please enter a valid end date.')
                elif end_date < start_date:
                    print('End month cannot be earlier than start month. Please enter a valid end date.')
                else:
                    break
            except:
                print('Invalid date format. Please enter a valid end date.')
    
    
    return ticker_option, tickers, frequency, start, end

# %%
def forex_monthly(ticker_option, tickers, start, end):
    
    '''This function takes in tickers, start date, and end date, and generates an Excel file with month-end data for each ticker.'''
    
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
        
        # get url of the last day
        url = get_url(last_day)
        
        # get dataframe of the last day
        df = get_df(url)
        
        # filter the df with date, and then get the last row of each ticker
        monthly_close_all = df[df['Date']==np.datetime64(last_day)].groupby(['Ticker']).last().reset_index()
        
        # if user chose to retrieve all currency pairs
        if ticker_option == 'a':
            # use monthly_close_all as final output without handling missing values on non-trading days 
            df_list.append(monthly_close_all)
        
        # if user specified currency pairs or uploaded a list of tickers
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
                            # print the message
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
                
                # if the dataframe is still empty, then append a row with NaN
                if ticker_monthly_close.empty:
                    ticker_monthly_close = pd.concat([ticker_monthly_close, pd.DataFrame([[ticker, np.datetime64(last_day), np.nan]], columns=ticker_monthly_close.columns)])
                
                # append the dataframe of each month to the list
                df_list.append(ticker_monthly_close)      
    
        # move to the next month
        start_date = start_date + relativedelta(months=1)
        
    # concatenate all dataframes in the list
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
    
    '''This function takes in tickers, start date, and end date, and generates an Excel file with daily data for each ticker.'''
    
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
    
    '''This function is the main function that calls other functions.'''
    
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
# run the main function
forexite()
