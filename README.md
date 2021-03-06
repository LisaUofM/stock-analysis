# VBA Module 2 Challenge

## Refactoring Stocks Analysis

### Overview of Project
The VBA module challenge features a new client, Steve. Steve has requested an analysis of green stocks, their volumes and their returns. We have used VBA code to create indexes, loops, conditionals and formatting to produce a succinct view of green stocks. https://github.com/LisaUofM/stock-analysis/blob/main/green_stocks.xlsm. 

Steve's next request is to expand the dataset to include entire stock market, so we have subsequently refactored the code for scaling the analysis to a larger population of stocks. https://github.com/LisaUofM/stock-analysis/blob/main/VBA_Challenge_v2.xlsm

#### Purpose
The purpose of this project is to refactor VBA Code of a stock analysis and demonstrate its effectiveness in reducing runtimes to prepare for larger data sets. 

### Results 

#### Original Analysis 

A macro for the initial analysis was created under subroutine "AllStocksAnalysis" in the file **green_stocks**. The runtime for retrieving stocks data for 2017 was 0.4960938 seconds. The runtime for retrieving stocks data for 2018 was 0.5039062 seconds. 

![Runtime for 2017 stock analysis](https://github.com/LisaUofM/stock-analysis/issues/2#issue-1092894322)

![Runtime for 2018 stock analysis](https://github.com/LisaUofM/stock-analysis/issues/1#issue-1092893889)


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
A macro with the refactored analysis was created under subroutine "AllStocksAnalsysRefactored" in the file **Vba_Challenge**. The runtime for retrieving stocks data for 2017 was 0.06640625, increasing runtime speed by **87%** (0.4961-0.06640625)/0.4961. The runtime for retrieving stock data for 2018 was 0.0675 seconds, increasing runtime speed by **87%**(0.5039062-0 0675)/0.5039062. 

![Runtime for refactored 2017 analysis](https://github.com/LisaUofM/stock-analysis/issues/6#issue-1097032060)

![Runtime for refactored 2018 analysis](https://github.com/LisaUofM/stock-analysis/issues/7#issue-1097032340)

The key differences between the original and refactored code are the use of a tickerIndex and the definition of tickerIndex variables as arrays (tickerVolumes(12), tickerStartingPrices(12) and tickerEndingPrices(12) as arrays. Using tickerIndex to find, store and return these variables resulted in reduced runtimes of 87% and 87% mentioned in the paragraph above. 

There are four key changes to the original code (1) created a tickerIndex, (2) set the variables as Arrays and (3) created a for loop to set the tickerVolumes to 0(4) and another for loop to retrieve the tickerVolumes, tickerStartingPrices and the tickerEndingPrices. No nested loop is used. 


(1) Creation of a ticker index from the array of tickers used in step 2 of the original analysis.  
```
tickerIndex = 0

```
(2) Creation of three output arrays for the Volumes, StartingPrices and EndingPrices variables. (Compare with step three of the original analysis.)
```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```
(3) Created a for loop to set the ticker Volumes to 0. 
```
For j = 0 To 11
  tickerVolumes(j) = 0
Next j

```
(4) Created another for loop to retrieve the tickerVolumes, tickerStartingPrices and the tickerEndingPrices. 

```
For i = 2 To RowCount
  
  If Cells(i, 1).Value = **tickerIndex** Then
        
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
       
  End If
        
  If Cells(i - 1, 1).Value <> **tickerIndex** And Cells(j, 1).Value = **tickerIndex** Then
            
    tickerStartingPrices(tickerIndex) = Cells(j, 6)
    
  End If
        
  If Cells(i + 1, 1).Value <> **tickerIndex** And Cells(j, 1).Value = **tickerIndex** Then
        
  tickerEndingPrices(tickerIndex) = Cells(j, 6)
            
  End If
    
    Next i
```

#### 

### Summary 

#### Advantages and Disadvantages of Refactoring Code in general
An advantage of refactoring code is that it reduces runtimes by making the code more efficient. If there are several processes in a batch, and the stocks analysis is one process, refactoring is an advantage, especially if the dataset is expected to increase. For example, if we were to run the refactored code on stocks in the Russell 2000, we would save 70 and 71 seconds for each year respectively. https://github.com/LisaUofM/stock-analysis/issues/8#issue-1097034914

If the dataset is small and the only batch process (like Steve's analysis of 12 stocks), the orignal code could be used without issue. Setting up another subroutine or fixing an existing subroutine for minimal rows of data can be labor intensive and unneccessary.  

#### Advantages and Disadvantages of the original and refactored VBA script

Advantages are that the runtimes are significantly reduced so the script is more capable of running a larger dataset. Disadvantages are that, while the script became more efficient in producing an output, it required more keystrokes in the coding process. For example, the "ticker" variable became "tickerIndex," the "volumes" became the "tickerVolumes," the "startingPrices" became the "tickerStartingPrices" etc.  Additional keystrokes make the code more vulnerable to human error and, as a result, could make the debugging process more time-consuming.  

