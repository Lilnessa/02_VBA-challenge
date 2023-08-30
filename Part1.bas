Attribute VB_Name = "Module4"
Sub tickersymbol_ws()

Set ws = ActiveSheet

For Each ws In Worksheets

'   add column titles
    ws.Range("J1").Value = "Ticker Symbol"
    ws.Range("k1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("m1").Value = "Total Stock Volumn"
  
    ws.Columns("j:Q").AutoFit
    
Dim ticker As String

Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'    MsgBox (lastrow)

'Part 1: ticker symbol
    For i = 2 To lastrow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            ws.Range("J" & Summary_Table_Row).Value = ticker
            Summary_Table_Row = Summary_Table_Row + 1
        End If

    Next i
Next ws

End Sub

