Attribute VB_Name = "Module7"
Sub totalvol_ws()

Set ws = ActiveSheet

For Each ws In Worksheets

Dim stockvolumn As Double
    stockvolumn = 0

Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'    MsgBox (lastrow)
    
'   Part 4: The total stock volume of the stock.
    
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            stockvolumn = stockvolumn + ws.Cells(i, 7).Value
            ws.Range("M" & Summary_Table_Row).Value = stockvolumn

            Summary_Table_Row = Summary_Table_Row + 1
            stockvolumn = 0

        Else
            stockvolumn = stockvolumn + ws.Cells(i, 7).Value

        End If

    Next i

'Format cells: Total stock volumn number includes commas

     ws.Range("M:M").NumberFormat = "#,###"

Next ws

End Sub
