Sub alphabets()
For Each ws In Worksheets
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"
ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"
Dim Ticker As String
Dim change As Double
Dim Volume As Double
Dim Opens As Double
Dim Closes As Double
ws.Range("N2").Value = ws.Range("C2").Value
Volume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
For i = 2 To LastRow


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker = ws.Cells(i, 1).Value
Opens = ws.Cells(i + 1, 3).Value
Closes = ws.Cells(i, 6).Value
Volume = Volume + ws.Cells(i, 7).Value
ws.Range("J" & Summary_Table_Row).Value = Ticker
ws.Range("M" & Summary_Table_Row).Value = Volume
ws.Range("N" & Summary_Table_Row + 1).Value = Opens
ws.Range("O" & Summary_Table_Row).Value = Closes
change = ws.Range("O" & Summary_Table_Row).Value - ws.Range("N" & Summary_Table_Row).Value
ws.Range("K" & Summary_Table_Row).Value = change
Summary_Table_Row = Summary_Table_Row + 1
Volume = 0
Opens = 0
change = 0
Else
Volume = Volume + ws.Cells(i, 7).Value
End If
Next i

LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
For i = 2 To LastRow
If ws.Cells(i, 14).Value <> 0 Then
percentage_change = (ws.Cells(i, 15).Value - ws.Cells(i, 14).Value) / ws.Cells(i, 14).Value
ws.Cells(i, 12).Value = percentage_change
Else
ws.Cells(i, 12).Value = "n/a"
End If
Next i

For i = 2 To LastRow
If (ws.Cells(i, 11).Value > 0) Then
    ws.Cells(i, 11).Interior.ColorIndex = 4
      Else
        ws.Cells(i, 11).Interior.ColorIndex = 3
     End If
     Next i
     
ws.Range("R2") = Application.WorksheetFunction.Max(ws.Range("L:L"))
ws.Range("R3") = Application.WorksheetFunction.Min(ws.Range("L:L"))
ws.Range("R4") = Application.WorksheetFunction.Max(ws.Range("M:M"))
For i = 2 To LastRow
If ws.Range("R2") = ws.Cells(i, 12).Value Then
IncreaseTicker = ws.Cells(i, 10).Value
ws.Range("q2") = IncreaseTicker
End If
If ws.Range("R3") = ws.Cells(i, 12).Value Then
DecreaseTicker = ws.Cells(i, 10).Value
ws.Range("q3") = DecreaseTicker
End If
If ws.Range("R4") = ws.Cells(i, 13).Value Then
VolumeTicker = ws.Cells(i, 10).Value
ws.Range("Q4") = VolumeTicker
End If
Next i
ws.Columns("L:L").NumberFormat = "0.00%"
ws.Range("R2:R3").NumberFormat = "0.00%"
ws.Columns(14).ClearContents
ws.Columns(15).ClearContents
Next ws
End Sub

