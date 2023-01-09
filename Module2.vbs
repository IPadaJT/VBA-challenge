Sub hw()
'Coded with help of Instructor Eli
    Dim lastrow As Long
    Dim i As Long
    Dim ticker As String
    Dim opening As Double
    Dim closing As Double
    Dim volume As Double
    Dim percentchng As Double
    
For Each ws In Worksheets
    opening = ws.Cells(2, 3).Value
    ticker = ws.Cells(2, 1).Value
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Count = 2
    volume = 0
    opening = ws.Cells(2, 3).Value
    ticker = ws.Cells(2, 1).Value
    
    'This labels new column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'This labels the functionality rows/columns
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    
      
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'When (i,1) no longer matches (i+1,1) then it will sum up the volumes up to that point
         closing = ws.Cells(i, 6).Value
         volume = ws.Cells(i, 7).Value + volume
         
         percentchng = (closing - opening) / opening
         
         ws.Cells(Count, 9).Value = ticker
         ws.Cells(Count, 10).Value = closing - opening
         
         'This section is for condtional formatting of the yearly change column
         If (closing - opening > 0) Then
            ws.Cells(Count, 10).Interior.ColorIndex = 4
            ElseIf (closing - opening < 0) Then
                ws.Cells(Count, 10).Interior.ColorIndex = 3
            End If
                
         ws.Cells(Count, 11).Value = percentchng
         ws.Cells(Count, 11).NumberFormat = "0.00%"
         
        'This section is for condtional formatting of the percent change column
         If percentchng > 0 Then
            ws.Cells(Count, 11).Interior.ColorIndex = 4
            ElseIf percentchng < 0 Then
                ws.Cells(Count, 11).Interior.ColorIndex = 3
            End If
        
         ws.Cells(Count, 12).Value = volume
         
         'MsgBox (ticker & ":" & opening & ":" & closing & ":" & Percent) Checks to see if ticker is counting to changes
         
         opening = ws.Cells(i + 1, 3).Value
         ticker = ws.Cells(i + 1, 1).Value
         Count = Count + 1
         volume = 0
    End If
    
    Next i
    Next ws
End Sub