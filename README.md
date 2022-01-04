# VBA Module 2 Challenge

## Refactoring Stocks Analysis

### Overview of Project
The VBA module challenge features a new client, Steve. Steve has requested an analysis of green stocks, their volumes and their returns. We have used VBA code to create indexes, loops, conditionals and formatting to produce a succinct view of green stocks. https://github.com/LisaUofM/stock-analysis/blob/main/green_stocks.xlsm. 

Steve's next request is to expand the dataset to include entire stock market, so we have subsequently refactored the code for scaling the analysis to a larger population of stocks.https://github.com/LisaUofM/stock-analysis/blob/main/VBA_Challenge.xlsm 

#### Purpose
The purpose of this project is to refactor VBA Code of a stock analysis and demonstrate its effectiveness in reducing runtimes to prepare for larger data sets. 

### Results 

#### Original Analysis 

A macro for the initial analysis was created under subroutine "AllStocksAnalysis" in the file **green_stocks**. The runtime for retrieving stocks data for 2017 was 0.4960938 seconds. The runtime for retrieving stocks data for 2018 was .5039062 seconds. 

![Runtime for 2017 stock analysis](https://github.com/LisaUofM/stock-analysis/issues/2)

![Runtime for 2018 stock analysis](https://github.com/LisaUofM/stock-analysis/issues/1)


The original code included six key steps (summarized to compare the original and refactored code): 

(1) Formatting the output sheet with a header row that includes Ticker, Daily Volume and Return 
```
Cells(3,1).Value = "Ticker"
Cells(3,2).Value = "Total Daily Volume" 
Cells(3,3).Value = "Return" 
```
(2) Initializing an array of tickers
```
Dim tickers(11) As String
       
tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"
```
(3) Preparing the data for analysis 
```
Dim startingPrice As Single
Dim endingPrice As Single

Sheets(yearValue.Activate)

RowCount = Cells(Rows.Count,"A").End(xlUp).Row
```
(4) Looping through the tickers using a "for" loop and setting totalVolume to 0 
```
For i = 0 to 11
ticker=tickers(i)
totalVolume = 0 
```

(5) Creating a nested loop to loop through the rows in the data to find the total volume, the starting price and the ending price using If statements and, lastly, closing the nested loop.

```
Worksheets(yearValue).Activate
For j = 2 To RowCount

If Cells(j, 1).Value = ticker Then
          
 totalVolume = totalVolume + Cells(j, 8).Value
                
End If

If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
         
 startingPrice = Cells(j, 6).Value
            
End If
If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
 endingPrice = Cells(j, 6).Value
                   
End If

Next j
```
(6) Outputing the data for each of the 12 stock tickers and closing the first loop. 
```
Worksheets("All Stocks Analysis").Activate
  Cells(4 + i, 1).Value = ticker
  Cells(4 + i, 2).Value = totalVolume
  Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
           
Next i
```

#### Refactored Analysis 
A macro with the refactored analysis was created under subroutine "AllStocksAnalsysRefactored" in the file **Vba_Challenge**. The runtime for retrieving stocks data for 2017 was 0.4375 seconds, increasing runtime speed by **12%** (0.4961-0.4375)/0.4961. The runtime for retrieving stock data for 2018 was 0.4335938 seconds, increasing runtime speed by **14%**(0.5039062-0.4335938)/0.5039062. 



#### 

### Summary 

#### Advantages and Disadvantages of Refactoring Code in general 

#### Advantages and Disadvantages of the original and refactored VBA script

