# forexite_data_reader

This Python script can be used to retrieve historical forex and cryptocurrency closing prices from a broker's website: https://www.forexite.com/traderoom/.  
Users can define tickers, time periods, and frequency of the data.  
The script will generate an Excel file to store the retrieved and cleaned data.  
The script handles the missing values on the non-trading days by linear interpolation (when retrieving daily data) or by retrieving the most recent value (when retrieving the month-end data).
