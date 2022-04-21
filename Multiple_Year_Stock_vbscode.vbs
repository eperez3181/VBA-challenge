Sub Census_Data()

For Each ws In Worksheets

Dim i As Long
Dim ticker As String
Dim vol_total As Double
vol_total = 0
Dim table_rows As Integer
table_rows = 2
Dim start As Double
Dim last As Double



ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

ws.Columns("K").NumberFormat = "0.00%"


For i = 2 To lastrow
    ticker = ws.Cells(i, 1).Value
    vol_total = vol_total + ws.Cells(i, 7).Value
    
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        start = ws.Cells(i, 3).Value
        
    End If
    
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
    ticker = ws.Cells(i, 1).Value
    
    last = ws.Cells(i, 6).Value
    
    ws.Range("I" & table_rows).Value = ticker
    
    ws.Range("J" & table_rows).Value = last - start
    
    ws.Range("K" & table_rows).Value = (last - start) / start
    
    ws.Range("L" & table_rows).Value = vol_total
    
    table_rows = table_rows + 1
    
    vol_total = 0
    
    End If


Next i

    lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

For i = 2 To lastrow

    If ws.Cells(i, 10).Value >= 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
    
    End If
    
    If ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
    
    End If

    
Next i

Next ws

End Sub
