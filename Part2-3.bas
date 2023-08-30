Attribute VB_Name = "Module5"
Sub yearlychange_ws()

Set ws = ActiveSheet

For Each ws In Worksheets

Dim yearlyopen As Double
Dim yearlyend As Double
Dim yearlychange As Double
Dim percentchange As Double
    percentchange = 0
Dim stockvolumn As Double
    stockvolumn = 0

Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'    MsgBox (lastrow)

'Part 2: Yearly change from the opening price at the beginning of a
'           given year to the closing price at the end of that year.
    
    For i = 2 To lastrow
'
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            yearlyopen = ws.Cells(i, 3).Value
                
    '        ws.Range("n" & Summary_Table_Row).Value = yearlyopen
    
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            yearlyend = ws.Cells(i, 6).Value
    '        ws.Range("o" & Summary_Table_Row).Value = yearlyend
    
    '        Summary_Table_Row = Summary_Table_Row + 1
        End If
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            yearlychange = yearlyend - yearlyopen
            ws.Range("k" & Summary_Table_Row).Value = yearlychange
    
'            Summary_Table_Row = Summary_Table_Row + 1
        End If
    



'Part 3: The percentage change from the opening price at the beginning
'           of a given year to the closing price at the end of that year.
   
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            percentchange = yearlychange / yearlyopen
            ws.Range("L" & Summary_Table_Row).Value = percentchange
            
            Summary_Table_Row = Summary_Table_Row + 1
        End If
        
    Next i
    
'Format ws.Cells
'    1) Yearly Change negative=red and positive=green

        lrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

            For j = 2 To lrow

                If ws.Cells(j, 11).Value > 0 Then
                    ws.Cells(j, 11).Interior.ColorIndex = 43
                ElseIf ws.Cells(j, 11).Value < 1 Then
                    ws.Cells(j, 11).Interior.ColorIndex = 3
                End If
            Next j

'    2)Percent change is percent number

     ws.Range("L:L").NumberFormat = "0.00%"

'    3)Total stock volumn number includes commas

     ws.Range("M:M").NumberFormat = "#,###"


Next ws

End Sub
