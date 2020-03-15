Sub challenge():
For Each ws In Worksheets

    Dim maxin As Double
    Dim minrow As Double
    Dim maxtotalvolume As Double
    Dim newlastrow As Double
    
    newlastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    'MsgBox (newlastrow)
    
        maxin = WorksheetFunction.Max(ws.Range("k2:k" & newlastrow))
        maxde = WorksheetFunction.Min(ws.Range("k2:k" & newlastrow))
        maxtotalvolume = WorksheetFunction.Max(ws.Range("L2:L" & newlastrow))

            'MsgBox (maxtotalvolume)
            
        maxrow = WorksheetFunction.Match(maxin, ws.Range("K2", "k" & newlastrow), 0)
        minrow = WorksheetFunction.Match(maxde, ws.Range("K2", "k" & newlastrow), 0)
        maxtotalrow = WorksheetFunction.Match(maxtotalvolume, ws.Range("L2", "L" & newlastrow), 0)
        
    
        
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % decrease"
    ws.Range("O4") = "Greatest Total Volume"

    
    ws.Cells(2, 16).Value = ws.Cells(maxrow + 1, 9).Value
    ws.Cells(2, 17).Value = maxin
    ws.Cells(2, 17).NumberFormat = "0.00%"

    ws.Cells(3, 16).Value = ws.Cells(minrow + 1, 9).Value
    ws.Cells(3, 17).Value = maxde
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ws.Cells(4, 16).Value = ws.Cells(maxtotalrow + 1, 9).Value
    ws.Cells(4, 17).Value = maxtotalvolume
    ws.Cells(4, 17).NumberFormat = "#,###,###"
    
Next ws

End Sub
