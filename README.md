# Stock Analysis through VBA 
## Overview of Project
For the purpose of this project, I worked on analyzing stock performance through large datasets for the years 2017 and 2018. As part of the analysis, I generated the total daily volume and overall return of each stock ticker. However, my goal was to look at larger datasets in the future, and in order to do so, I needed to refactor my original code to have the analysis run faster and more efficiently. I worked on changing the original code and set timers to compare the speed of the refactored code to the original. 
## Results
#### Stock Performance
When it came to stock performance, the years 2017 and 2018 showcased a dramatic difference. In 2017, we see almost entirely positive returns, while in 2018, the returns were almost entirely negative. 2017 saw DQ at an almost 200% return rate while in 2018, DQ had a -62% return. 
![2017_Stock_Performance](https://github.com/noorsami12/stock-analysis/blob/04d17a8f0d16238a597838a8f63bd9327ee79168/stock%20analysis%202017.png)
![2018_Stock_Performance](https://github.com/noorsami12/stock-analysis/blob/04d17a8f0d16238a597838a8f63bd9327ee79168/stock%20analysis%202018.png)
#### Execution Times
The execution time of the refactored code is significantly decreased. The original code ran in 0.59 seconds for 2017, while the refactored code ran in 0.09 seconds for 2017. For 2018, the original code was also 0.59 seconds while the refactored code was slightly less at 0.08 seconds. Overall, the refactored code was much faster than the original code. 
![OG_2017_Runtime](https://github.com/noorsami12/stock-analysis/blob/04d17a8f0d16238a597838a8f63bd9327ee79168/og%20code%202017%20time.png)
![2017_Runtime](https://github.com/noorsami12/stock-analysis/blob/04d17a8f0d16238a597838a8f63bd9327ee79168/Resources/VBA_Challenge_2017.png)
![OG_2018_Runtime](https://github.com/noorsami12/stock-analysis/blob/04d17a8f0d16238a597838a8f63bd9327ee79168/og%20code%202018%20time.png)
![2018_Runtime](https://github.com/noorsami12/stock-analysis/blob/04d17a8f0d16238a597838a8f63bd9327ee79168/Resources/VBA_Challenge_2018.png)
The main area of difference in the refactored code was creating a tickerIndex that would connect across all four arrays used in the code to make the processes more efficient. 
```
'1a) Create a ticker Index
   
    Dim tickerIndex
    tickerIndex = 0


    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerIndex = 0
    Next i
     
    ''2b) Loop over all the rows in the spreadsheet.

        For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
     
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         
         ElseIf Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
    
            tickerIndex = tickerIndex + 1
            
         End If
    
        
         Next i
```
## Summary
####Refactoring Code – Pros and Cons
Refactoring code can be incredibly advantageous when it comes to improving the speed and efficiency of a program. However, refactoring code can also be time-consuming and cause complications amidst the code that may not have been there before. For example, I ran into multiple instances of conflict between different variable types throughout refactoring, which was not a problem I faced while writing the original code. It would be easy to break the code and ruin the data because of issues like this. When it comes to refactoring code, we need to weigh the pros and cons and make sure that the end result is worth it. In this case, an almost .5 second speed decrease in the runtime seems worth it to me! 
####Original vs. Refactored VBA script
The original VBA script I wrote was overall shorter and quicker to write. It contained less variables and might be less confusing for another programmer to look at and work with. The refactored VBA script runs more efficiently and will be an asset when it comes to analyzing larger datasets. However, the code itself was longer and more time-consuming and may not be as intuitive to the outside eye. 
