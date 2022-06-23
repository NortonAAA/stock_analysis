# An Analysis of Green Stocks for Parents
June 22, 2022

## Overview of Project
This project is to build a tool for analysis of a group of Green stocks for investment by parents.

### Purpose
The purpose of this project was to build a tool inside an Excel spreadsheet to compare two years of performance between 2017 and 2018. A VBA macro was built to built out a report for each year and format according to performance for that year.

## Summary
### Recommendations of Stock Investments
After building out the VBA code, the results for the data set show that 2017 was a good year from all but 1 stock (TERP). It is important to look at 2018, however, as that highlights the volatility of this sector. There are two investments that have really strong returns and appear to be worth some investment. These two are ENPH: Enphase Energy, Inc. and RUN: SunRun. Particulary ENPH was at least double-digit growth over the two years.

![](https://github.com/NortonAAA/stock_analysis/blob/main/VBA_Challenge_2017.png) ![](https://github.com/NortonAAA/stock_analysis/blob/main/VBA_Challenge_2018.png)

### Code Comparison
There were two ways to approach this code and one was more beneficial in processing time. This came looking for opportunity to refactor the code to make this run faster and with less memory. Kely call outs as highlight below in the refactor image below is [1] Changing from Double to Single for Starting and Ending Price variables should reduce the amount of storage needed. [2] Utilization the tickerIndex array allows for a reduction of a Nested Loop cycle and able to be held and pulled into the output easier. [3] Formatting is integrated into the sub routine instead of having another button to call the formatting.
#### Original Code vs. Refactored Code
<img src="https://github.com/NortonAAA/stock_analysis/blob/main/AllStockAnalysis_Original.png" height="50%" width="45%" > <img src="https://github.com/NortonAAA/stock_analysis/blob/main/AllStockAnalysis_Refactored.png" height="50%" width="45%" >
### Performance of Code
Based on the above changes the ultimately performance of the code ran almost 7% faster than the original code
![](https://github.com/NortonAAA/stock_analysis/blob/main/Code_Performance.png)






## Summary
