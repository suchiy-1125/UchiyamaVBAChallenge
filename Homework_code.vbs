Sub HW():

    For Each ws In Worksheets
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percentage Change"
    ws.Range("L1") = "Total Volume"
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim k As Double
    k = 2
    
    Dim startPrice As Double, endPrice As Double
    startPrice = ws.Cells(2, 6).Value
    

    Dim volumetotal As LongLong
    
    volumetotal = 0
    
    For i = 2 To lastrow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(k, 9) = ws.Cells(i, 1).Value
            
            
            If startPrice <> 0 Then
                endPrice = ws.Cells(i, 6).Value
                    
                ws.Cells(k, 10) = endPrice - startPrice
                ws.Cells(k, 11) = (endPrice - startPrice) / startPrice
            
                ws.Cells(k, 11).NumberFormat = "0.00%"
                If Cells(k, 10) > 0 Then
                
                ws.Cells(k, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(k, 10).Interior.ColorIndex = 3

                End If
                
                                
            End If
            volumetotal = volumetotal + ws.Cells(i, 7).Value
            ws.Cells(k, 12) = volumetotal
            ws.Cells(k, 12).NumberFormat = "#,###,###"
            
        
        k = k + 1

        startPrice = ws.Cells(i + 1, 6).Value
    
        volumetotal = 0
        
        Else
            volumetotal = volumetotal + ws.Cells(i, 7).Value
            
        End If
                    
    Next i

    Next ws

End Sub
