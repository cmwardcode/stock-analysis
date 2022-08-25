# An Analysis of Stocks Through Refactoring
## Overview of Project
Excel VBA Developer was used to program an Excel workbook to analyze stock dataset. After the initial code was written it underwent refactoring to loop through data once and decrease run time. 
### Purpose and Background
Provided with a stock dataset, client requested a user-friendly analysis be developed and ran at the click of a button. The dataset provided included two years of data, 2017 and 2018. Each sheet included 13 tickers, date, open, high, low, close, adj close and volume.  A simple test macro was written and ran. Upon successfully running testing the macro and analysis sheet was then created, “DQ Analysis” for which the client wanted to focus on this ticker and find out the total volume and return for each year that could be ran by clicking a button and then selecting a year. A second sheet was added to analyze all stock tickers, “All Stock Analysis.” On this sheet a code was developed to analyze the total volume for each of the other 12 tickers and the return. The return was color coded to visualize a positive or negative return. The sheet could be run by a button and entering a year to analyze. After creating this analysis, the code was then refactored to analyze the “All Stocks Analysis” in less time with its own run button.
## Analysis and Results 
### Analysis
This analysis will focus on comparing the refactored code to the original to determine whether the refactoring successfully made the VBA script run faster. The refactored code which can be referenced below was formatted to output all results to the “All Stocks Analysis” sheet and output the data into the data into the correct headers, “Ticker”, “Total Daily Volume”, and “Return”. An array of the tickers was initialized, and the sheet was activated. Code was then written to collect a row count to et the number of rows to loop over. A ticker index variable was then added and set to zero. The variable was created to access the correct index across four different arrays. Three output arrays were incorporated: ticker volume, ticker starting price, and ticker ending price. The loop was initialized to set the ticker volumes to zero. A second loop was designed to iterate over all the rows of the spreadsheet. Inside the second loop the script was designed to increase the volume of the ticker volume output variable and add the ticker volume for each stock ticker using the ticker index variable. This loop used if-then statements to check if the current row is the first row with the selected ticker index and then assign starting price to the starting price variable. A second if-then statement was inserted to check if the current row was the last row with the selected index and if it is to then assign the current closing price to the ending price variable. The script then increased the ticker index to do the same for each stock ticker. A third loop went through the arrays to output the data into the visible headers on the sheet: “Ticker,” “Total Daily Volume,” and “Return.”
```VBA
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
       tickerVolumes(i) = 0

    Next i  
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End if
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value 
            
            End if
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


```
### Results
The results were that the refactored code did run faster than the original. The improvement in speed was an average of a 14% increase for 2017 stock data and 16% increase for 2018. The images of run time for the refactored code are shown below.

![VBA_Challenge_2017]( https://github.com/cmwardcode/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018]( https://github.com/cmwardcode/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary
### Advantages
The advantage of refactoring the code is ultimately creating a better code and to decrease the run time of the VBA script.  Other advantages are a cleaner code reduces complexity for other users understanding. It is easier to maintain and requires less time to update and support which can lower costs. 
### Disadvantages
The largest disadvantage is the extra time it takes to refactor. Refactoring may also not offer a significant decrease in run time. If refactoring is not performed correctly, it could possibly introduce new bugs and errors into the code. The benefits of refactoring may not be self-evident since the functionality stays the same so the client may not notice. 
