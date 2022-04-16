# Stock Analyisis

## Project Overview

Steve parents are passionate about green energy stocks, and decided to invest all their money into DAQO New Energy Corp (DQ). Steve asked us to look into DQ performance for the years of 2017 and 2018. Steve also wanted to analyze other stocks in the same industry, compare it with DQ, and explore the option of diversifyng his parent's investment.
I used Excel VBA code to perform this analysis; however, the code needs to be refactored in order to make it run more effficiently.

### Purpose

The purpose of this project is to refactor an Excel VBA code that analyzes the annual return of twelve different stock investments. By refactoring it, I make the code more efficient by using less memory, and I make it easier for others to understand the code. 

## Refactoring the Code

The refactored code ran faster beacause I created arrays instead of having nested loops. I achieved this by setting the tickerIndex variable equal to zero, and then creating three more output arrays for tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

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

Three more `For` loops with `If-Then` statements were created to loop over the rows, and check if the current row matches the first or last row of the selected ticker to store the values for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

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
        
 Finally, the code looped through the arrays to output the Total Volume and Return in the spreadsheet.
 
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
        
          Worksheets("All Stocks Analysis").Activate
          Cells(4 + i, 1).Value = tickers(i)
          Cells(4 + i, 2).Value = tickerVolumes(i)
          Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
          
        Next i
        
The formatting section of the code remained the same.

## Results

While both sets of code show the same results for the return of each stock, the original code ran in 0.76 seconds for 2017, and in 0.67 seconds for 2018.
The refactored code ran in 0.16 seconds for 2017, and in 0.15 seconds for 2018. Please see screenshots below:


![](/Resources/VBA_Challenge_2017.PNG)     ![](/Resources/VBA_Challenge_2018.PNG)


## Summary

1. What are the advantages and disadvantages of refactoring code?

   - Advantages: the code is more efficient, meaning that it takes less time to perform a routine. Additionally, it could be easier for others to understand the refactored code as it is usually better structured.
   - Disadvantages: It can be hard to work with someone else's code, specially if you don't understand it. It can also be time consuming, and you put at risk the outcome of the project.
  
2. How do these pros and cons apply to refactoring the original VBA script?

   The original VBS script works correctly, but it is almost five times slower than the refactored one. This could be a huge disadvantage when working with larger codes or larger files. The refactored code was faster and its structure was easier to follow. I belive that if someone else needs to modify or add to the code, it would be easier to work with the refactored one. 


