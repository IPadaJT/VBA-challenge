Sub minfunc()
Dim minValue As Double
Dim minTicker As String
Dim lastrow As Long
Dim i As Long


Count = 2

For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
    If ws.Cells(i, 11).Value <= minValue Then
        minValue = ws.Cells(i, 11).Value
minTicker = ws.Cells(i, 9).Value
End If


Next i

'MsgBox (maxValue)
'MsgBox (maxTicker)
ws.Cells(3, 17).Value = minValue
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = minTicker

Next ws

End Sub