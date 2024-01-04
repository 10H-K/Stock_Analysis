# Utilizing VBA Scripting to Analyze Multiple Years of Stock Data

## Overview ##

The objective of this project is to create a Visual Basic for Applications (VBA) script in Microsoft Excel that loops through three consecutive years of data for approximately 3,000 stocks and outputs summarizing information for each individual stock. The initial stock data is organized into five main categories: open price, close price, highest intraday price, lowest intraday price, and volume traded. This information is recorded for each individual stock on every single trading day throughout the course of a calendar year. By using VBA, this data is analyzed to determine the annual variation of each individual stock, reported as both a monetary change and corresponding percent change, along with the total volume of each stock that was traded during the course of the year. This condensed data is further analyzed to ascertain the specific stocks that achieved the greatest percent increase, greatest percent decrease, and greatest total volume traded during each calendar year. The application of VBA in this situation allowed for all of the data processing to be automated, preventing the need for a repetitive approach to the more than 750,000 rows of information contained in each worksheet.

## Process ##

The stock data that served as the basis of this project is organized into an excel workbook titled "Multiple Year Stock Data," which contains a seperate worksheet for each calendar year (2018, 2019, and 2020) under consideration. This particular arrangement necessitated a VBA script that would loop through each worksheet in the workbook and accomplish both levels of data analysis. To achieve this goal, four separate Sub procedures were created in VBA:

1) Sub StockRawDataAnalysis()
    * This Sub procedure will summarize the initial raw stock data and output the following information:
        * Yearly Change
            *  Monetary difference between open price at the beginning of the calendar year under consideration and close price at the end of that year for each individual stock.
        * Yearly Percent Change
            * Conversion of the monetary Yearly Change into a percentage change. 
        * Total Stock Volume
            * Cumulative sum of the total volume traded for each individual stock throughout the course of the calendar year under consideration.
    * Basic structure of Sub procedure:
        1) For Loop allows Sub procedure to run through every worksheet in the Multiple Year Stock Data workbook.
        2) DataOutput variable is defined.
            * Directs all condensed stock information determined by this Sub procedure to a specific location in the worksheet.
        3) Second For Loop analyzes each row of raw stock data one at a time, where:
            * Variables StockTicker and TotalVolume are defined.
                * StockTicker will assign each unique stock ticker in DataOutput.
                * TotalVolume will determine the cumulative annual total volume traded for each individual stock.
            * Variable OpenPrice is defined.
                * Value of OpenPrice is stock price at the beginning of the calendar year under consideration.
                * OpenPrice is set once the stock ticker identity in the previous row of data is not the same as the current row of data.
            * Change in stock ticker identity is precisely identified.
                * Determined when the stock ticker identity in the next row of data is not the same as the current row of data.
                * After the stock ticker identity change has been established, these steps will follow:
                    * StockTicker is set before the identity changes in the following row of data.
                    * ClosePrice variable is defined
                        * Assigned its corresponding value in the final row of data before the stock ticker identity changes.
                    * PriceChange and PercentChange variables are defined.
                        * Assigned values through formulaic manipulation of the ClosePrice and OpenPrice variables.
                            * Ultimately giving the values for Yearly Price Change and Yearly Percent Change.
                    * Volume traded in final row of data is added to the running total held by TotalVolume.
                        * TotalVolume now has the final value for the total stock volume traded. 
                    * Summarized data that has been obtained will be printed to the DataOutput location.
                    * All of the variables will be reset in preparation for stock ticker identity change.
                    * Row will be added to the DataOutput location in preparation for stock ticker identity change.
                * If there is no change in stock ticker identity, then the stock volume traded will continually be added to TotalVolume as the loop progresses.
        4) The For Loop will end on final row of raw stock data and this process will start over again in the next worksheet.               

2) Sub StockCondensedDataAnalysis()
    * This Sub procedure will evaluate the summarized stock data obtained in the previous Sub procedure and output the following highlighting information:
        * Greatest Percent Increase
            * Stock ticker that has achieved the largest positive Yearly Percent Change amongst all stocks during the calendar year under consideration.
        * Greatest Percent Decrease
            * Stock ticker that has achieved the largest negative Yearly Percent Change amongst all stocks during the calendar year under consideration.
        * Greatest Total Volume
            * Stock ticker that has achieved the largest volume of stock traded throughout the course of the calendar year under consideration.
    * Basic structure of Sub procedure:
        1) For Loop allows Sub procedure to run through every worksheet in the Multiple Year Stock Data workbook.
        2) DataOutput variable is defined.
            * Directs all stock highlight information determined by this Sub procedure to a specific location in the worksheet.
        3) Variables used to determine Greatest Percent Increase / Greatest Percent Decrease / Greatest Total Volume are defined and subsequently reset in preparation for the upcoming For Loop.
        4) Variables used to determine the stock ticker identities that correspond to the Greatest Percent Increase / Greatest Percent Decrease / Greatest Total Volume are defined and subsequently reset in preparation for the upcoming For Loop.
        5) For Loop created so that each row of summarized stock data will be evaluated one at a time, where:
            * Variables CurrentPercent, CurrentVolume, and CurrentTicker will hold values for the Yearly Percent Change, Total Stock Volume, and stock ticker identity in the specific row of summarized stock data that the For Loop is evaluating.
                * By sequentially holding these three data points in every single row of summarized stock data, it will be possible to continously search for the Greatest Percent Increase / Greatest Percent Decrease / Greatest Total Volume one row at a time.
            * Variables GreatestIncrease, GreatestDecrease, and GreatestVolume will start storing values as the For Loop progresses. The variables StockTickerIncrease, StockTickerDecrease, StockTickerVolume will start storing the stock ticker identities that correspond to the values stored in GreatestIncrease, GreatestDecrease, and GreatestVolume.
                *  Whenever CurrentPercent is greater than GreatestIncrease, then GreatestIncrease will store CurrentPercent as its new value. Also, the variable StockTickerIncrease will store CurrentTicker as its new stock ticker identity.
            * Whenever CurrentPercent is less than GreatestDecrease, then GreatestDecrease will store CurrentPercent as its new value. Also, the variable StockTickerDecrease will store CurrentTicker as its new stock ticker identity.
            * Whenever CurrentVolume is greater than GreatestVolume, then GreatestVolume will store CurrentVolume as its new value. Also, the variable StockTickerVolume will store CurrentTicker as its new stock ticker identity.
         6) After all rows of the summarized stock data have been analyzed, the updated and finalized variables from step 5 will be printed to the DataOutput location.
         7) This process will start over again in the next worksheet.

3) Sub NamingAndFormatting()
    * This Sub procedure will format all of the data found in each worksheet so that the overall workbook looks visually appealing.
    * Basic structure of Sub procedure:
        1) For Loop allows Sub procedure to run through every worksheet in the Multiple Year Stock Data workbook.
        2) Assignment of names to specific cells, which will descibe and organize the data obtained from the two previous Sub procedures.
        3) Variables CurrencyColumn, PercentageColumnA, and PercentageColumnB are defined.
            * Used later to format the data obtained from the two previous Sub procedures.
        4) Insertion of a table to organize the stock highlight information obtained in the previous Sub procedure.
        5) Color formatting for the Yearly Change and Yearly Percent Change
            * For Loop created so that each row of Yearly Change and Yearly Percent Change will be evaluated one at a time, where:
                * Green color will fill the Yearly Change and Yearly Percent Change cells when both values are greater than zero.
                * Red color will fill the Yearly Change and Yearly Percent Change cells when both values are less than zero.
                * Yellow color will fill the Yearly Change and Yearly Percent Change cells when both values are equal to zero.
                * Column for Yearly Change is formatted as a currency and column for Yearly Percent Change is formatted as a percentage.
        6) The cells for Greatest Percent Increase and Greatest Percent Decrease are formatted as percentages.
        7) The font, font size, column width, and text alignment are formatted.
        8) This process will start over again in the next worksheet.

4) Sub MasterSwitch()
    * This Sub procedure will make using the previous three Sub procedures a simple and efficient process.
    * Basic structure of Sub procedure:
        1) Calls the Sub procedure StockRawDataAnalysis
        2) Calls the Sub procedure StockCondensedDataAnalysis
        3) Calls the Sub procedure NamingAndFormatting
    * This format allows for the entire data organization process to be completed by running the single macro MasterSwitch.

## Results ##

