Sub stockhw()



Dim ws As Worksheet
Dim ticker As String
Dim total As Double
total = 0
Dim minvalue As Single
Dim maxvalue As Single
Dim change As Single
Dim pchange As Single

Dim summary As Single

On Error Resume Next


For Each ws In Worksheets
ws.Range("A:L").Columns.AutoFit
summary = 2
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row



For i = 2 To lrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker = ws.Cells(i, 1).Value
    total = total + ws.Cells(i, 7).Value
    ws.Range("I" & summary).Value = ticker
    ws.Range("L" & summary).Value = total
    summary = summary + 1
    total = 0
    Else
    total = total + ws.Cells(i, 7).Value
    End If
    
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value And ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
    minvalue = ws.Cells(i, 3).Value
    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value And ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value And ws.Cells(i, 1) <> 0 Then
    maxvalue = ws.Cells(i, 6).Value
    ws.Range("J" & summary - 1).Value = maxvalue - minvalue
        If ws.Range("J" & summary - 1).Value > 0 Then
        ws.Range("J" & summary - 1).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & summary - 1).Value <= 0 Then
        ws.Range("J" & summary - 1).Interior.ColorIndex = 3
        End If
    ws.Range("K" & summary - 1).Value = (maxvalue / minvalue) - 1
    ws.Range("K" & summary - 1).Style = "Percent"
    End If
    
    

Next i


Next ws



End Sub