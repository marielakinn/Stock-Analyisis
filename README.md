# Stock-Analyisis

## Project Overview

Steve parents are passionate about green energy stocks, and decided to invest all their money into DAQO New Energy Corp (DQ). Steve asked us to look into DQ performance for years 2017 and 2018, and to analyze other stocks in the same industry to compare it with DQ, and to explore the option of diversify their investment.
I used Excel VBA code to perform this analysis, but now I want to refactor it to make the code run more effficient.

### Purpose

The purpose of this project is to refactor an Excel VBA code that analyzes the annual return of twelve different stok investments. By refactoring I make the code more efficient by using less memory, and I make it easier for others to understand the code. 

## Refactoring the code

The refactored code ran faster beacause I created arrays instead of having nested loops. I achieved this by setting the tickerIndex variable equal to zero, and then creating three output arrays for tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

    '1a) Create a ticker Index
          tickerIndex = 0

    '1b) Create three output arrays
          Dim tickerVolumes(12) As Long
          Dim tickerStartingPrices(12) As Single
          Dim tickerEndingPrices(12) As Single
          
Then, I created a `For` loop to initialize all the variables to zero.

    '2a) Create a for loop to initialize the tickerVolumes to zero.
          For i = 0 To 11
              tickerVolumes(i) = 0
              tickerStartingPrices(i) = 0
              tickerEndingPrices(i) = 0
          Next i

Three more `For` loops with `If-Then` statements were created to loop over the rows, and check if the current row matches the first or last row of the selected ticker to set their starting and ending price.

    '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
           'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        'End If
         End If
            
        
    '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
         End If
            
    '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
        
 Finally, I looped through the arrays to output the Total Volume and Return in the spreadsheet.
 
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
        
          Worksheets("All Stocks Analysis").Activate
          Cells(4 + i, 1).Value = tickers(i)
          Cells(4 + i, 2).Value = tickerVolumes(i)
          Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
          
        Next i
        
The formatting of the code remained the same.

## Results

While both sets of code show the same results for the retrn of each stocks, the original code ran in 0.76 seconds for 2017, and in 0.67 seconds for 2018.
The refactored code ran in 0.16 seconds for 2017, and in 0.15 seconds for 2018. Please see screenshots below:




## Summary

1. What are the advantages and disadvntages of refactoring code?
2. How do these pros and cons apply to refactoring the original VBA script
