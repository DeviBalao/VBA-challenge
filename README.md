# VBA-challenge
To upload files related to the VBA-challenge - VBA scripting to analyze generated stock market data.

Steps to follow to execute the VBA challenge code to analyze the stock data:
================================================================================
1. Download and open the excel file "Multiple_year_stock_data" available in Module 2 challenge in Canvas.
2. In the first sheet (2018), select the Developer tab and insert a button (preferably in column M,N so processed data is not hidden behind the button). 
3. In Assign Macro popup, edit Button1_Click() to ProcessData_Click() and click New.
4. Copy the code from "Analyze_Stocks_Code.txt" and paste it in Code Editor in Module1. 

	Note: while copying, paste the code inside Sub ProcessData_Click() to Sub ProcessData_Click(). The rest of the code (starting from "Sub AnalyzeStocks(ws)") can be pasted below it.

5. In the excel, modify the display text from Button 1 to Process Data. This is optional.
7. Click the button(Process Data / Button 1) in the first sheet (2018). Processing will take a couple of minutes.
8. Data will be processed and the following changes will be done in all 3 sheets (2018,2019,2020) of the excel file:

  ------------------------------------------------------------------------------------------------------------------------------------------------
	a) 4 columns - Ticker, Yearly Change, Percent Change and Total Stock volume will be displayed with unique Ticker and its corresponding values.

	b) Next to these 4 columns, the stock (Ticker) with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume" will be displayed. 

	c) Yearly Change and Percent Change will be in red for negative values and green for positive values.

  --------------------------------------------------------------------------------------------------------------------------------------------------

** Output screenshots are available in Output_Screeshots folder.

===========================================================================================================================================================
