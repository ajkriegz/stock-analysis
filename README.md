# VBA Stock Analysis
## Overview of project

### Purpose of this analysis
Steve just graduated with a finance degree and has taken on his parents as clients. Steve's parents want to invest in green energy but have done very little research into stocks. Steve wants to investigate stock data in order to diversify their funds and has gathered stock data in several green energy companies. This analysis uses VBA to analyze and visualize this data in Microsoft Excel to inform and guide Steve as he works with his parents to invest in the stock market.

## Results
Below are the two tables of stocks with daily volume and rate of return for the years 2017 and 2018, respectively.

![Stocks, 2017](https://github.com/ajkriegz/stock-analysis/blob/main/resources/2017_stock_analysis.png)

![Stocks, 2018](https://github.com/ajkriegz/stock-analysis/blob/main/resources/2018_stock_analysis.png)

2017 offered a much better rate of return on the year for this range of stocks than 2018 with few exceptions. While showing lower yearly returns than the previous year, Enphase Energy Inc (ENPH) finished 2018 quite strong, and Sunrun Inc (RUN) posted large gains for the year.

The rate of return for the year was calculated by first finding the starting and ending prices for the year using the following the code:
```
If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
```

These numbers were then used to calculate the quotient of the ending price for a given ticker and its starting price, then turning that number into a percentage, as shown by `Cells(4 + y, 3).Value = tickerEndingPrices(y) / tickerStartingPrices(y) - 1`
---

Below are the refactored script's execution times:

![VBA refactored script run time for 2017](https://github.com/ajkriegz/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

![VBA refactored script run time for 2018](https://github.com/ajkriegz/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

Compare this to the previous execution times:

![Original script execution time, 2017](https://github.com/ajkriegz/stock-analysis/blob/main/resources/VBA_original_script_2017.png)

![Original script execution time, 2018](https://github.com/ajkriegz/stock-analysis/blob/main/resources/VBA_original_script_2018.png)

The refactored script is almost ten times faster than the original. For twelve stocks, the impact is minimal. However, circumstances may change. The code is written to be easily edited for as many stocks as desired in the future. Additionally, older machines may experience increased processing time and may struggle with too large a task.

## Summary

### The advantages and disadvantages of refactoring code

Refactoring code can be a largely beneficial process. It improves program efficiency, restructures poorly designed code structure, improves readability and cohesion, and offers a chance to reduce the number of bugs in the code. Additionally, it provides an excellent opportunity for peer review.

However, it is not always ideal to refactor code. If enough time has passed, it will be more difficult to remember why or how a section of code works. The developer is also spending additional time and energy on something that is already working, and may in fact introduce more bugs to the program if too little time is allotted for testing.
---

In the case of this analysis, the subroutine's efficiency was increased and may reduce large amounts of processing time in the future. However, for the current scope of stocks, saving less than a second off program execution may be estimated to be a waste of resources in a manager's eyes.

