Attribute VB_Name = "Module2"
Sub functionality()
    ' Greatest percent increase, decrease, and greatest total volume
    ' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets
    
        ' Add Header titles
        ' Greatest Percent Increase
        ws.Cells(2, 15).Value = "Greatest_perc_increase"
        ws.Cells(3, 15).Value = "Greatest_perc_decrease"
        ws.Cells(4, 15).Value = "Greatest_total_volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Fill in the values
        ' last row of columns K and L
        LastRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
        initial_inc = 0
        initial_dec = 0
        initial_vol = 0
        
        For k = 2 To LastRow
            'Finding greatest percent increase value
            If ws.Cells(k, 11).Value > 0 Then
                If ws.Cells(k, 11).Value > initial_inc Then
                    initial_inc = ws.Cells(k, 11).Value
                    inc_ticker = ws.Cells(k, 9).Value
                Else
                    initial_inc = initial_inc
                End If
            
            'Finding greatest percent decrease value
            ElseIf ws.Cells(k, 11).Value < 0 Then
                If ws.Cells(k, 11).Value < initial_dec Then
                    initial_dec = ws.Cells(k, 11).Value
                    dec_ticker = ws.Cells(k, 9).Value
                Else
                    initial_dec = initial_dec
                End If
            End If
            
            'Finding greatest total volume
            If ws.Cells(k, 12).Value > initial_vol Then
                initial_vol = ws.Cells(k, 12).Value
                vol_ticker = ws.Cells(k, 9).Value
            Else
                initial_vol = initial_vol
            End If
            
        Next k
        
        ' Fill in with the values found
        ws.Cells(2, 16).Value = inc_ticker
        ws.Cells(2, 17).Value = initial_inc
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = dec_ticker
        ws.Cells(3, 17).Value = initial_dec
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = vol_ticker
        ws.Cells(4, 17).Value = initial_vol
        
    Next ws
    
End Sub


