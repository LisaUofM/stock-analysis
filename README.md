# stock-analysis
Module 2: VBA assignment
# VBA Module 2 Challenge

## Refactoring Stocks Analysis

### Overview of Project
The VBA module challenge features a new client, Steve. Steve has requested an analysis of green stocks, their volumes and their returns. We have used VBA code to create indexes, loops, conditionals and formatting to produce a succinct view of green stocks. https://github.com/LisaUofM/stock-analysis/blob/main/green_stocks.xlsm. 

Steve's next request is to expand the dataset to include entire stock market, so we have subsequently refactored the code for scaling the analysis to a larger population of stocks. 

#### Purpose
The purpose of this project is to demonstrate how we have refactored the VBA Code and its effectiveness in reducing runtimes in preparation for larger data sets. 

### Results 

#### Original Analysis 

A macro for the initial analysis was created under subroutine "AllStocksAnalysis" in the file **green_stocks**. The runtime for retrieving stocks data for 2017 was 0.4960938 seconds. The runtime for retrieving stocks data for 2018 was .5039062 seconds. 

The original code included six primary steps (summarized to show differences between the original and refactored code): 

(1) Formatting the output sheet with a header row that includes Ticker, Daily Volume and Return 
```
Cells(3,1).Value = "Ticker"
Cells(3,2).Value = "Total Daily Volume" 
Cells(3,3).Value = "Return" 
```
(2) Initializing an array of tickers 

(3) Preparing the data for analysis 

(4) Looping through the tickers using a "for" loop 

(5) Creating a nested loop to loop through the rows in the data to find and store the total volume, the starting price and the ending price using If statements 

(6) Outputing the data for each of the 12 stock tickers. 

#### Refactored Analysis 
A macro with the refactored analysis was created under subroutine "AllStocksAnalsysRefactored" in the file **Vba_Challenge**. The runtime for retrieving stocks data for 2017 was 0.4375 seconds, increasing runtime speed by **12%** (0.4961-0.4375)/0.4961. The runtime for retrieving stock data for 2018 was 0.4335938 seconds, increasing runtime speed by **14%**(0.5039062-0.4335938)/0.5039062. 

#### 

### Summary 

#### Advantages and Disadvantages of Refactoring Code in general 

#### Advantages and Disadvantages of the original and refactored VBA script

