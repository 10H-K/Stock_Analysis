# Utilizing VBA Scripting to Analyze Multiple Years of Stock Data

#Overview

The objective of this project was to create a Visual Basic for Applications (VBA) script in Microsoft Excel that loops through three consecutive years of data for approximately 3,000 stocks and outputs summarizing information for each individual stock. The initial stock data was organized into five main categories: open price, close price, highest intraday price, lowest intraday price, and volume traded. This information was recorded for each individual stock for every single trading day throughout the course of a calendar year. Through the use of VBA, this data was analyzed to determine the annual variation of each individual stock, which was reported as both a monetary change and a corresponding percent change, along with the total volume of each stock that was traded during the course of the year. This condensed data was then further analyzed to ascertain the specific stocks that achieved the greatest percent increase, greatest percent decrease, and greatest total volume traded during each calendar year. The application of VBA in this situation allowed for all of the data processing to be automated, preventing the need for a repetitive approach to the more than 750,000 rows of information contained in each worksheet.

#Purpose

The stock data that served as the basis of this project was organized into an excel workbook titled "Multiple Year Stock Data," which contained a seperate worksheet for each calendar year (2018, 2019, and 2020) that was considered. This particular arrangement necessitated a VBA script that would loop through each worksheet in the workbook and accomplish both levels of data analysis. To accomplish this task, three separate Sub procedures were created:

1) Sub StockRawDataAnalysis()
    * The main purpose of this Sub procedure is to summarize the initial raw stock data and output the following information: Yearly Change in monetary terms, Yearly Percent Change based on the monetary change, and Total Stock Volume traded.
        * Yearly Change is determined by calculating the difference between the opening price at the beginning of a given year and the closing price at the end of that year for each individual stock.
        * Yearly Percent Change is determined by converting the currency value of the Yearly Change into a percentage.
        * Total Stock Volume traded is determined by summing together all of the volume traded totals throughout the course of a calendar year for each individual stock.
    * The basic structure of this Sub procedure is:
        * First, a For loop allows this this Sub procedure to be run through every worksheet in the workbook.
        * Next, the variables for Stock Ticker Identity and 
        * Next, another For loop is set up in order to go through the raw stock data one row at a time where:
            * The unique Stock Ticker identity
            * Two 
            * 
      
4) 
5) 
