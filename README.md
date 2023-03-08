# Module-2---Challenge-VBA
Module 2 Challenge as part of Monash University Data Analytics Bootcamp

# Stock Analysis
This repository contains a VBA script that analyses stocks across the 2018, 2019, and 2020 financial years across each separate worksheet in an Excel workbook. The script loops through all the stocks on a year-by-year basis and outputs the following summary table in column J, K, L, M, P, Q, and R on the relevant year's worksheet. The summary table contains the following information:

The unique ticker symbol for each stock, held within each individual financial year.
Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year for each specific stock (ticker).
> The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year for each specific stock (ticker).
> The total stock volume for each specific stock (ticker).
> Additionally, the script also returns the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume" for each specific financial year.
> Noting that again that this is run across all worksheets within the workbook.
 
## How to Use
•      Download the Excel workbook that contains the stock data you want to analyse, ensuring the data is separated by each financial year.
•      Once open, enable the Developer tab in the Ribbon.
•      In the Developer tab, click on Visual Basic to open the Visual Basic Editor.
•      In the Visual Basic Editor, right-click on the workbook name and select Insert > Module.
•      Copy and paste the VBA script from this repository into the new module.
•      Press F5 to run the script.
•      The script will output the analysis results for each specific year in the worksheet relevant to that year across columns J, K, L, M, P, Q, and R. This will occur across every worksheet within the workbook.
 
## Notes
•      The script assumes that the stock data is arranged in the following columns: ticker symbol, date, opening price, high price, low price, closing price, and total stock volume.
•      The script also assumes that the stock data is arranged by year, with each year in a separate worksheet.
•      If your stock data is arranged differently, you may need to modify the VBA script to fit your needs.
