Sub StockCycle()
'Declare variable for all worksheets
 Dim ws As Worksheet
 'Insert header labels for ticker, yearly change, percent change, and total stock volume
Cells(1, "H").Value = "Ticker"
Cells(1, "I").Value = "Yearly Change"
Cells(1, "J").Value = "Percent Change"
Cells(1, "K").Value = "Total Stock Volume"
'Create a variable to hold the ticker symbol
 Dim TickerSymbol As String
'Set a variable for holding the total stock symbol
Dim Volume_Total As Double
Volume_Total = 0
'Create a variable to hold the opening price for the beginning of the year
Dim OpenPrice As Double
'Create a variable to hold the closing price for the end of the year
Dim ClosingPrice As Double
'Create a variable to hold the yearly change
Dim YearlyChange As Double
 'Create a variable to hold the percent change
Dim PercentChange As Double
'Keep track of the location for each ticker in column i and establish which row to start the loop
Dim TickerRow As Integer
TickerRow = 2
'Create for each statement to loop through all worksheets
For Each ws In Worksheets
    'Set a variable for the last row of the worksheet
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    ' Loop through all ticker symbols
    For i = 2 To LastRow
        ' Check if we are still within the same ticker symbol.
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Find the ticker symbol value
            TickerSymbol = ws.Cells(i, 1).Value
            'Find the Volume total value
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            'Find the Year Close value
            ClosingPrice = ws.Cells(i, 6).Value
            'Find the Yearly Change value
            YearlyChange = ClosingPrice - OpenPrice
            'Find the Percent Change value
            If ClosingPrice > 0 Then
                PercentChange = YearlyChange / ClosingPrice
            End If
            'Print the above information into the correct columns
            ws.Range("H" & TickerRow).Value = TickerSymbol
            ws.Range("K" & TickerRow).Value = Volume_Total
            ws.Range("I" & TickerRow).Value = YearlyChange
            ws.Range("J" & TickerRow).Value = PercentChange
            'Change column J to a percent
            ws.Range("J" & TickerRow).NumberFormat = "0.00%"
            'MsgBox (TickerSymbol)
            'Loop through by adding one to the ticker row
            TickerRow = TickerRow + 1
            'Reset the volume total
            Volume_Total = 0
            OpenPrice = 0
            'MsgBox (Volume_Total)
        ' If the cell immediately following a row is the same ticker symbol...
        Else
            If (OpenPrice = 0) Then
                OpenPrice = Cells(i, 3).Value
            End If
            ' Add to the volume total
            Volume_Total = Volume_Total + Cells(i, 7).Value
        End If
        If ws.Cells(i, 9).Value > 0 Then
            ws.Cells(i, 9).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 9).Interior.ColorIndex = 3
        End If
    Next i
Next ws

End Sub
