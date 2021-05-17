# VBA-Challenge
##
Overview of Project: 
Steve has asked me to analyze a green stock for his parents to determine wherther it worth investing in. To do the analysis, i needed to determine the daily volume and annual return on the stocks. After anylyzing one stock, i continued to look through 11 more green stocks, to give steve a few options to choose from

##
Results: 
####New Code Below
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
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
    Worksheets(yearValue).Activate
    'get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    tickerIndex = 0
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    'initialize ticker volumes to zero
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    For i = 2 To RowCount
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
    Next i
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
            Cells(i, 3).Interior.ColorIndex = 4
        Else
            Cells(i, 3).Interior.ColorIndex = 3
        End If
    Next i
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
End Sub

###Old Code



##
Summary: 
What are the advantages or disadvantages of refactoring code?
Advantages of refactoring code is that you do not have to start from scratch. You take an existing piece of code that work, and try and make it better. The advantage here is saving time and not having to re-invent the wheel, plus if you can make it more efficient, it will be less taxing on the resources. 
The disadvanteage here is you could potentially break a code which was working before, and by the time you figure out where the issue is or how the code is written, you find that it might have been easier or faster to just write the code from scratch

##
How do these pros and cons apply to refactoring the original VBA script?
The pros here was i didn't have to write the code from scratch. The con is that i broke something and it took me several hours to try and figure it out. Probably would have been faster to write it from scratch
