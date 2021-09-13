# Green Energy Stock Analysis - Refactored

## Project Overview
The aim of this project is to provide an analysis of a group of green energy stocks (potential investment targets). The completed analysis was then refactored to improve speed by obtaining all relevant data from one loop through the larger dataset, rather than looping through it for each target stock.

## Results
Our original completed analysis obtained the correct information, but there was significant room to improve the efficiency of the code. The dataset contains 3012 lines for each calendar year, one line describing the activity for a single stock for each market day. The data is sorted alphabetically by ticker, and then by date. 
### Original Analysis
The original code created a loop through the entire dataset for each individual stock. 
``` 
For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
       Worksheets(yearValue).Activate
        For j = 2 To RowCount
            If Cells(j, 1).Value = ticker Then
            'increase totalVolume by value in current row
                totalVolume = totalVolume + Cells(j, 8).Value
     
            End If
        
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
        
            End If

            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value

            End If
        
        Next j
```
This run took 0.29 seconds for 2017 and 0.28 seconds for 2018.

### Refactored Analysis
Given how the data is presented though, it is possible to loop through the rows once, extracting and storing the needed data into arrays. 
```
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
        
    Next tickerIndex
    
    tickerIndex = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        ticker = tickers(tickerIndex)
        
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = ticker Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
       
        If Cells(i, 1).Value = ticker And Cells(i - 1, 1).Value <> ticker Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                

           '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
   
    Next i
```

Then we can present it in an appropriate table. 

```'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
              
        
    Next i
```

![Screenshot_2017 Results](/Resources/Table_2017results.png)

### Refactored Results
Run time is reduced by 77% in each year.

![Refactored run time 2017](/Resources/VBA_Challenge_2017.png)
![Refactored run time 2018](/Resources/VBA_Challenge_2018.png)

## Summary
### Pros and Cons of Refactoring
Refactoring is a useful tool in many projects. It serves to clean up code, making it run more efficiently and making it more readable and easier to work with for other developers in the future. If you have a project which you have worked through step by step, taking a big picture view may allow you to organize it in a better way. The code itself produces the same result, but does so in a more elegent, efficient way.
The downsides of refactoring are primarily to do with time. You need to work in small pieces to make sure that you are not introducing problems into a code that was, after all, working properly to begin with. If a code is written for a very specific task, and there is very little chance that the code will be used for any other purpose, or in any different way, it may not be worth the investment of time to go back and restructure it.

### Original vs. Refactored Script in this project.
In this project, the refactored script certainly provides a more elegant approach to the problem. With the small number of stocks (12), the original script still completes the analysis in .26 seconds. While the percentage decrease in time with the refactored code is very large, the difference is barely noticable to the end user. 
However, the refactored script could easily be scaled up to assess many more stocks than the current 12. As you increase the number of stocks, and therefore the number of rows, the time difference will become more and more noticable. It is always better to be prepared for what you might do next with your project, and refactoring this code does that. 
