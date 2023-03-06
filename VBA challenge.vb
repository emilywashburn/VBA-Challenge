Sub ticker()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

Dim LastRow As Long
Dim i As Long
Dim j As Long
Dim counter As Long

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17) = "Value"

j = 2
counter = 2

' search for the last row in column A
LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

' loop through the rows
For i = 2 To LastRow

'check to see if next cell is different than current
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
' place value in ticker column
    ws.Cells(counter, 9) = ws.Cells(i, 1).Value
    'calculate yearly change
    ws.Cells(counter, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
    'set the color of the yearly change cells
    If ws.Cells(counter, 10).Value < 0 Then
        ws.Cells(counter, 10).Interior.ColorIndex = 3
        Else: ws.Cells(counter, 10).Interior.ColorIndex = 4
    End If
    'calculate the percentage
    ws.Cells(counter, 11).Value = (ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value
    
    'calculate the total volume by using the sum function
    ws.Cells(counter, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
'increment counter and j
counter = counter + 1
j = i + 1
    
End If

Next i

lastrow2 = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
increase = 0
decrease = 0
volume = 0


For i = 2 To lastrow2

If ws.Cells(i, 11).Value > increase Then
    increase = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).Value = increase
End If
If ws.Cells(i, 11).Value < decrease Then
    decrease = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 17).Value = decrease
End If
If ws.Cells(i, 12).Value > volume Then
    volume = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = volume
End If

Next i

Next ws

End Sub
