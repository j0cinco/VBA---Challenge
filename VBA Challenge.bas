Attribute VB_Name = "Module1"
Sub summary():
'setting up wsheet
For Each ws In ThisWorkbook.Worksheets
Dim sheetname As String

'columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'rows
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'loop
tickernum = 2

j = 2

lastrow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To lastrow1


'IF FUNCTIONS
'Year Change
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ws.Cells(tickernum, 9).Value = ws.Cells(i, 1).Value
ws.Cells(tickernum, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

        'color
        If ws.Cells(tickernum, 10).Value > 0 Then
        ws.Cells(tickernum, 10).Interior.ColorIndex = 4
        Else
        ws.Cells(tickernum, 10).Interior.ColorIndex = 3
        End If
        
'Percent Change
If ws.Cells(j, 3).Value <> 0 Then
PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
ws.Cells(tickernum, 11).Value = Format(PercentChange, "Percent")
Else
ws.Cells(tickernum, 11).Value = Format(0, "Percent")
End If

'Total Volume
ws.Cells(tickernum, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
tickernum = tickernum + 1
j = i + 1

End If

Next i

'Loop for new table
lastrow9 = ws.Cells(Rows.Count, 9).End(xlUp).Row

Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim HighestStock As Double

GreatestIncrease = ws.Cells(2, 11).Value
GreatestDecrease = ws.Cells(2, 11).Value
HighestStock = ws.Cells(2, 12).Value

For i = 2 To lastrow9
    
    'Greatest Increase
    If ws.Cells(i, 11).Value > GreatestIncrease Then
    GreatestIncrease = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
    Else
    GreatestIncrease = GreatestIncrease
    End If
    
    'Greatest Decrease
    If ws.Cells(i, 11).Value < GreatestDecrease Then
    GreatestDecrease = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
    Else
    GreatestDecrease = GreatestDecrease
    End If
    
    'Highest Stock
    If ws.Cells(i, 12).Value > HighestStock Then
    HighestStock = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
    Else
    HighestStock = HighestStock
    End If
    
Range("Q2:Q3").NumberFormat = "0.00%"


Next i


Next ws


End Sub
