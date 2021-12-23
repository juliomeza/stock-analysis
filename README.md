# Stock Analysis with VBA

## Overview of Project
### Purpose
The purpose of this project is to help Steve's parents make the best decision about which stocks to buy.

## Results
To simplified the analysis, I have filtered out all the stocks that performed below average.

### Total Daily Volume Above Average
1. Only 3 stocks have a 'Total Daily Volume' above the average in both years (2017 and 2018).
2. Out of the 3 stocks, only 'RUN' shows a volume increase from 2017 to 2018.

<img src="https://github.com/juliomeza/stock-analysis/blob/main/resources/Over%20Average%20Volume.png" width="600">

### Return Above Average 
1. Only 2 stocks have a 'Return' above the average in both years (2017 and 2018).
2. From the 2 stocks with Return above average, only 'ENPH' has a positive return in both years (2017 and 2018).

<img src="https://github.com/juliomeza/stock-analysis/blob/main/resources/Over%20Average%20Return.png" width="600">

### Three Possible Stocks to Buy
1. The price tag trend for the top three stocks (RUN, ENPH, SEDG) look very similar.
2. A deepest investigation is needed to select the best option.

<img src="https://github.com/juliomeza/stock-analysis/blob/main/resources/Ticker%20RUN.png" width="600">

<img src="https://github.com/juliomeza/stock-analysis/blob/main/resources/Ticker%20ENPH.png" width="600">

<img src="https://github.com/juliomeza/stock-analysis/blob/main/resources/Ticker%20SEDG.png" width="600">

### Code Execution Time
The execution times have improved by 400% in the 2017 data set, and by 500% in the 2018 data set. This is because the refactored script loops over the entire data set only once, rather than as many tickers there are.

2017
<p float="left">
  <img src="https://github.com/juliomeza/stock-analysis/blob/main/resources/2017-0.PNG" width="300">
  <img src="https://github.com/juliomeza/stock-analysis/blob/main/resources/VBA_Challenge_2017.PNG" width="300">
</p>

2018
<p float="left">
<img src="https://github.com/juliomeza/stock-analysis/blob/main/resources/2018-0.PNG" width="300">
<img src="https://github.com/juliomeza/stock-analysis/blob/main/resources/VBA_Challenge_2018.PNG" width="300">
</p>

## Summary
- The advantage of refactoring the code is to reduce the running time. This allows us to analyze greater amount of data without sacrificing the waiting time. 
- The only disadvantage on this specific refactoring case if that the data needs to be sorted before running the code.
- One of the limitations of the data set is that the values are recorded by day, instead of by minute. Stocks prices changes a lot during the day.
- One of the challenges was to convert the yy-mm-dd format to yy/mm/dd format to be recognize by the pivot table. After spending a lot of time playing with the date formating option, the solution was as simple as replacing the '-' with '/' using the Find and Replace option