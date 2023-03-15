# forexite_data_reader

This Python script can be used to retrieve historical forex and cryptocurrency closing prices from a broker's website: https://www.forexite.com/traderoom/.  


We can get the raw data for a given day from an URL in this format: https://www.forexite.com/free_forex_quotes/YYYY/MM/DDMMYY.zip

For example: https://www.forexite.com/free_forex_quotes/2023/03/020323.zip

The raw data downloaded from this URL contains the closing prices for every minute of the day for all currency pairs, but we are only interested in the closing prices at the end of the day. Also, we can only get data for a specific day from the website, rather than a range of dates or a specific frequency. This script was written to address these two problems.

This script will allow the user to specify the tickers, time period, and frequency of the data. All the data retrieval and manipulation will be done automatically by this script. At the end, this script will generate an Excel file to store the cleaned data.

The libraries used in this script include: requests, pandas, datetime, dateutil.
