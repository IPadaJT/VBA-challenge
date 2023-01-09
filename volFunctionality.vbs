Sub volFunc()
Dim volValue As Variant
Dim volTicker As String
Dim lastrow As Long
Dim i As Long


Count = 2

For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
    If ws.Cells(i, 12).Value >= volValue Then
        volValue = ws.Cells(i, 12).Value
volTicker = ws.Cells(i, 9).Value
End If


Next i

'MsgBox (maxValue)
'MsgBox (maxTicker)
ws.Cells(4, 17).Value = volValue
ws.Cells(4, 16).Value = volTicker

Next ws

End Sub
