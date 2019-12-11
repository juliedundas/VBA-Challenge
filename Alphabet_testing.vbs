alhpabetical_testing.vbs

Sub StockCycle()

'Declare a variable for all worksheets
Dim ws As Worksheet

'Create for each statement to loop through all worksheets
For Each ws In Worksheets

'Set a variable for the last row of the worksheet
LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row


'Create a variable to hold the ticker symbol
Dim TickerSymbol As String

' Set a variable for holding the total stock symbol
  Dim Volume_Total As Double
  Volume_Total = 0

  ' Keep track of the location for each ticker in column i
  Dim TickerRow As Integer
  TickerRow = 2

  ' Loop through all ticker symbols
  For i = 2 To LastRow

' Check if we are still within the same ticker symbol.
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 
 ' Set the ticker symbol
      TickerSymbol = ws.Cells(i, 1).Value

      ' Add to the ticker Ttotal
      Volume_Total = Volume_Total + ws.Cells(i, 7).Value
      
       ' Print the ticker symbol in the i column
      ws.Range("H" & TickerRow).Value = TickerSymbol
    
     'MsgBox (TickerSymbol)
     'MsgBox (Cells(i, 9).Value)
     
      ' Print the volume total to the j column
      ws.Range("K" & TickerRow).Value = Volume_Total

      ' Add one to the ticker row
      TickerRow = TickerRow + 1
    
       ' Reset the volume Total
      Volume_Total = 0 + Cells(i, 7).Value
      
     'MsgBox (Volume_Total)
    
      
         ' If the cell immediately following a row is the same ticker symbol...
    Else

      ' Add to the volume total
      Volume_Total = Volume_Total + Cells(i, 7).Value

    End If


  Next i
  
  'Create a variable to hold the percent change
  Dim PercentChange As Double
  PercentChange = 0
  
  'Create a variable to hold the opening price for the beginning of the year
  Dim OpenPrice As Double

  
  'Create a variable to hold the closing price for the end of the year
  Dim ClosingPrice As Double
  
  'Loop through ticker symbols
  'For j = 2 To LastRow
  
  'Check if we are still within the same ticker symbol
  If Cells(i + 1, 2).Value <> Cells(i, 2).Value Then
  
'Assign value to the OpenPrice variable
OpenPrice = ws.Cells(i, 2).Value

'Assign value to the ClosingPrice
If Cells(i + 1, 6).Value <> Cells(i, 6).Value Then
ClosingPrice = ws.Cells(i, 6).Value

PercentChange = OpenPrice - ClosingPrice

'Print the PercentChange in column I
ws.Range("I" & TickerRow).Value = PercentChange

MsgBox PercentChange


End If

Next i

Next ws

End Sub



Sub AddHeader()

'Declare a variable for all worksheets
Dim ws As Worksheet

For Each ws In Worksheets

'Add a header for the ticker symbol
ws.Range("H1").Value = "Ticker"
ws.Range("I1").Value = "Yearly Change"
ws.Range("J1").Value = "Percent Change"
ws.Range("K1").Value = "Total Stock Volume"

Next ws

End Sub
