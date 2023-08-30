Attribute VB_Name = "Module8"
Sub greatest_ws()

Set ws = ActiveSheet

For Each ws In Worksheets

'Part 5: Add functionality to your script to return the stock with the
'       "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
    
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("P2").Value = "Greatest % increase"
    ws.Range("P3").Value = "Greatest % decrease"
    ws.Range("P4").Value = "Greatest total volume"

    ws.Columns("j:Q").AutoFit
    
lrow = Cells(Rows.Count, 10).End(xlUp).Row
'    MsgBox (lrow)
    
    For k = 2 To lrow
    
        ws.Range("R2").Value = WorksheetFunction.Max(ws.Range("L2:L" & k))
        ws.Range("R3").Value = WorksheetFunction.Min(ws.Range("L2:L" & k))
        ws.Range("R4").Value = WorksheetFunction.Max(ws.Range("M2:M" & k))
        
        ws.Cells(2, 17).Value = WorksheetFunction.Index(ws.Range("J2:J" & k), _
            WorksheetFunction.Match(ws.Range("R2").Value, ws.Range("L2:L" & k), 0))
        ws.Cells(3, 17).Value = WorksheetFunction.Index(ws.Range("J2:J" & k), _
            WorksheetFunction.Match(ws.Range("R3").Value, ws.Range("L2:L" & k), 0))
        ws.Cells(4, 17).Value = WorksheetFunction.Index(ws.Range("J2:J" & k), _
            WorksheetFunction.Match(ws.Range("R4").Value, ws.Range("M2:M" & k), 0))
    Next k
    
        ws.Range("R2:R3").NumberFormat = "0.00%"
        ws.Range("r4").NumberFormat = "#,###"
        ws.Columns("j:Q").AutoFit

Next ws

End Sub



