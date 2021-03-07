Sub stocks()

Application.ScreenUpdating = False

For Each ws In Worksheets


Dim i, j As Integer
Dim rowcount As Long
Dim total As Double

rowcount = ws.Cells(Rows.Count, 1).End(xlUp).Row


'ticker and total volume columns


total = 0
j = 2

For i = 2 To rowcount
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        total = total + ws.Cells(i, 7).Value
        ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(j, 12).Value = total
        j = j + 1
        total = 0
    Else
        total = total + ws.Cells(i, 7).Value
    End If
Next i



'puts closing prices for each stock in yearly change column

j = 2

For i = 2 To rowcount
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        ws.Cells(j, 10).Value = ws.Cells(i, 6).Value
        j = j + 1
    End If
Next i



'gets opening prices for each stock, subtracts from yearly change column
'calculates percent change and populates column

j = 2

For i = 2 To rowcount
    If ws.Cells(i, 1).Value = ws.Cells(j, 9).Value Then
        
        'yearly change calc
        ws.Cells(j, 10).Value = ws.Cells(j, 10).Value - ws.Cells(i, 3).Value
        
        'percent change calc
        If ws.Cells(i, 3).Value = 0 Then
            ws.Cells(j, 11).Value = "n/a"
        Else
            ws.Cells(j, 11).Value = ws.Cells(j, 10).Value / ws.Cells(i, 3).Value
        End If
        
        j = j + 1
    End If
Next i

'labels + conditional formatting

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly change"
ws.Range("K1").Value = "Percent change"
ws.Range("L1").Value = "Total volume"
ws.Range("K:K").NumberFormat = "0.00%"

For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i

Next ws

Application.ScreenUpdating = True

End Sub




