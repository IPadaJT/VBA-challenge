Sub MaxFunc()
Dim maxValue As Double
Dim maxTicker As String
Dim lastrow As Long
Dim i As Long


Count = 2

For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
    If ws.Cells(i, 11).Value >= maxValue Then
        maxValue = ws.Cells(i, 11).Value
maxTicker = ws.Cells(i, 9).Value
End If


Next i

'MsgBox (maxValue)
'MsgBox (maxTicker)
ws.Cells(2, 17).Value = maxValue
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(2, 16).Value = maxTicker

Next ws

End Sub