# Stock-analysis
## Overview
   View the comparison of different stocks from 2017 and 2018 to reach a better investment decision. The analysis was conducted based on the given data set , and the data set code was refactorerd to attain an efficient result in lesser time. 
 ## Results
 Before refactoring, starter code was downloaded and followed the steps to activate the new code. Find the refactored code and instructions below.
 Sub Allstocksanalysisrefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStocksAnalysis").Activate
    
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
    
    
    '2a) Initialize ticker volumes to zero
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    '2b) loop over all the rows
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("AllStocksAnalysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i) - 1)
 
    Next i
    
    'Formatting
    Worksheets("AllStocksAnalysis").Activate
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
## Summary
  #### Advantages and disadvantages of refactoring code
         Refactoring a code makes a code cleaner and tailored for the demand. It helps in efficiency in terms of debugging and faster programming. Refactoring makes data user friendly by making it easier to read and interpret.
         Disadvantage or a challenge refactoring a code happens when the data is large and an established code is already running.
  #### Advantages and disadvantages of the original and refactored VBA script
         Advantage of refactoring is the reduced time to run the macro. Refactored code took only one fourth the time of the original code. 
         ![image](https://user-images.githubusercontent.com/100480390/158081216-e89e8091-7a35-4135-8024-8cac2dd3b22e.png)
          ![image](https://user-images.githubusercontent.com/100480390/158081228-4bfbef7a-472c-40c3-9ed6-c6cf2627d3f7.png)
         Disadvantage of refactoring VBA code includes longer and confusing coding syntax  making it less user friendly.

   
