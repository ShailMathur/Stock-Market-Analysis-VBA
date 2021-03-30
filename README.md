# VBA-Challenge to make a Stock Market Analysis

# Background
A VBA scripting project to analyze Real Stock Market Data. This analysis used two data-points one is the Test Data which is used while developing the script, and the other is Stock Data which is the main data to run the script. The stock data or the main data has three sheets categorized yearly 2014, 2015, 2016.

# Solution
The script loop through all the stocks data once and the following information displayed.

1. The Ticker Symbol - The script will loop through all the stocks for one year and list each ticker symbol in column "I".

2. Yearly Change - The script will excute Yearly change from opening price at the beginning of a given year to the closing price at the end of that year. This value is added in column "J" for every Ticker symbol.

3. Percent Change - The script calculates a percent change from opening price at the beginning of a given year to the closing price at the end of that year. This value is added in column "K" for every Ticker symbol.

4. The Total Stock Volume - The scripts computes the total stock volume in column "L" for each ticker.

5. As a Bonus, the script outputs the "Greatest % increase", "Greatest % decrease" and "Greatest total volume" among all Ticker symbol for each year.
