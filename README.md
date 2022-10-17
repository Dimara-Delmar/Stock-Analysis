# VBA Stock Analysis

## Overview of Project

### Purpose
Steve has recently graduated with a degree in finance, and he wants to help his parents invest in stock. The dataset he is currently working with is designed to analyze a few dozen stocks, but he wants to be able to expand his dataset to analyze more. The purpose of this project is to refactor the original code so that it can loop through the dataset only once instead of multiple times per stock, and to see if refactoring the code in this way will allow it to run more efficiently.  

## Results

### Stock performance of 2017 and 2018
2017 vastly outperformed 2018 in terms of daily volume and yearly returns, the only exceptions being the two stocks ENPH and RUN.

<img width="224" alt="All_Stocks_2017" src="https://user-images.githubusercontent.com/108738297/196158809-2221c4bc-e046-4310-b03b-d879778d9b9b.png"> <img width="224" alt="All_Stocks_2018" src="https://user-images.githubusercontent.com/108738297/196158839-0b510437-bf39-4d05-b09f-3c7f8bd0ab75.png">

Judging by the outcomes of these tables, if Steve were to recommend any of the stocks in this dataset to his parents ENPH and RUN would be the safest investment for them to make. 

### Refactored Code Comparison 
The original `for` loop was structured in a way that looped over the dataset multiple times. It checks through each current ticker and loops back around once it ends and moves on to the next one:

    'Loop through the tickers
    
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
    'Loop through the rows
    
        Worksheets("2018").Activate
        For j = 2 To RowCount
        
    'Find total volume for the current ticker
        
        If Cells(j, 1).Value = ticker Then
        
            totalVolume = totalVolume + Cells(j, 8).Value
            
        End If
        
    'Find the starting price for the current ticker
    
        If Cells(j, 1) = ticker And Cells(j - 1, 1) <> ticker Then
        
            startingPrice = Cells(j, 6).Value
                        
        End If
        
    'Find the ending price for the current ticker
    
        If Cells(j, 1) = ticker And Cells(j + 1, 1) <> ticker Then
        
            endingPrice = Cells(j, 6).Value
            
        End If
    
    Next j
    
    'Output the data for the current ticker
    
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
    
    
    Next i

To change this so that the data only needs to loop one time, a <tickerIndex> variable was added and used to access four other arrays in the script: `tickers`, `tickerVolume`, `tickerStartingPrices`, and `tickerEndingPrices`:

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

Using `if-then` statements and the newly added arrays, a `for` loop is used to check for the first row (starting price), the last row (ending price), and to increase the `tickerIndex` if the next row’s ticker doesn’t match previous one: 

    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
        tickerVolumes(i) = 0

    Next i
        
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.

        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d) Increase the tickerIndex.

            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i

Finally, a `for` loop is used to loop though the arrays and display the output in the worksheet:

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

### Run Times Comparison
The original script had a run time of 0.548 seconds for the year 2017, and 0.545 seconds for the year 2018. 

<img width="238" alt="Original_Runtime_2017" src="https://user-images.githubusercontent.com/108738297/196158273-893de8b4-e265-4417-a5a5-c1aae5d61239.PNG"> <img width="237" alt="Original_Runtime_2018" src="https://user-images.githubusercontent.com/108738297/196158308-17d1fa6d-4aad-401b-9e4d-4253b519c58a.PNG">

The edited script had a run time of 0.108 seconds for the year of 2017, and 0.109 seconds for the year of 2018.

<img width="237" alt="Refactored_Runtime_2017" src="https://user-images.githubusercontent.com/108738297/196158355-634ff1eb-6808-4aac-9eab-49a0dc2b3d73.png"> <img width="232" alt="Refactored_Runtime_2018" src="https://user-images.githubusercontent.com/108738297/196158386-6c5abc85-e0f8-4338-87dc-34729a3a3ae2.png">

From this comparison, we can conclude that the refactored script had a meaningful improvement on run time performance. 

## Summary

### Advantages and Disadvantages of Refactored Code
One of the advantages that comes with refactoring code is that it provides a pre-existing framework to work off, and since it’s code that already works, we just need to change it. In this case, we can look at what the original code was is designed to do and determine how to edit it from there. This process saves us the trouble of having to start from scratch and can give some useful insight on what to do next.   

However, the disadvantages of working with pre-existing code is the possibility of running into issues with debugging and errors. Since we’re not starting from the beginning, it can be difficult to understand where exactly the code is going wrong- If it’s something that that was added, or if the new things are just conflicting with old things in the code. This can be an even bigger issue when working with code that isn’t your own, making it hard to understand what the original code might be trying to do and how it can be fixed.

When refactoring the original VBA script, I found that working with pre-existing code helped in the fact that the code provided was already working, and since I know how it was written, it gave me a good idea on what needed to change to make it run more smoothly. unfortunately, I still ran into issues when I tried to edit and add things to the code. There was one specific issue that kept popping up and I don’t know how it got there or how it can be fixed. Sometimes when I run the refactored code, it gives me an entirely different run time for both years than it originally did. For example, it will say “This code ran in 8.398 seconds for the year 2017” instead of the original 0.108 seconds. 
