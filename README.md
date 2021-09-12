# Green Energy Stock Analysis - Refactored

## Project Overview
The aim of this project is to provide an analysis of a group of green energy stocks (potential investment targets). The completed analysis was then refactored to improve speed by obtaining all relevant data from one loop through the larger dataset, rather than looping through it for each target stock.

## Results
Our original completed analysis obtained the correct information, but there was significant room to improve the efficiency of the code. The dataset contains 3012 lines for each calendar year, one line describing the activity for a single stock for each market day. The data is sorted alphabetically by ticker, and then by date. The original code created a loop through the entire dataset for each individual stock. 
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

With the refactored analysis the runtime was significantly improved

![Refactored run time 2017](/Resources/VBA_Challenge_2017.png)
![Refactored run time 2018](/Resources/VBA_Challenge_2018.png)

## Summary


