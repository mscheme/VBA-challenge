# VBA-challenge
VBA Stock Ticker Analysis

The VBA Stock Ticker Analysis is a homework assignment from a Data Analytics Bootcamp.
I have included the original instructions in the 'VBA Challenge Instructions.md' file.

There were two sets of data provided. The first was 'alphabetical_texting.xlsx'.
This file was used during development and testing of the macros.
The final macros were run against a larger Multiple_year_stock_data file that was too large to be included in this repository.

There are three major assumptions made:
1. The data is sorted in order by Stock Ticker (column A)
2. The data is then sorted by date (January to December)
3. Only a single year of data is on a single sheet

Given the assumptions are correct and the data provided, running the macros will produce a summary of data for each stock ticker and sheet.
The summary of data for each stock ticker includes: yearly change, percentage change, and total stock volume.
The summary of data for the sheet includes: greatest % increase, greatest % decrease, and greatest total volume.

To Run the Macros:
Either import the file 'VBA_Challenge_Macro.bas' or copy/paste the code from 'StockTickerChallenge_code.vbs' into a Visual Basic module.

There are three macros:
1. loopThruWorksheets - loops through every sheet in the open workbook
                     - calls StockTicker for every sheet
                     - calculates time taken to run
                     - displays a message to the user when the macros has finished running
2. StockTicker - performs all calculations (for yearly change, percentage change, etc.) in a sheet.
               - calls writeHeaders
3. writeHeaders - writes the new headers (Ticker, Yearly Change, etc.) in a sheet

If you want to test the macros on a single sheet, run 'StockTicker'.
If you want to run the macros on an entire workbook, run 'loopThruWoorkSheets'

Additionally, I have included screenshots of my results for each year of the Multiple_Year_Stock_Data file.
Mult_Stock_2014.png, Mult_Stock_2015.png, and Mult_Stock_2016.png
