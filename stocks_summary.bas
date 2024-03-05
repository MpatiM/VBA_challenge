Attribute VB_Name = "Module1"
Sub stocks_summary()

    ' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets
    
        ' Add Header titles
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly_Change"
        ws.Cells(1, 11).Value = "Percentage_Change"
        ws.Cells(1, 12).Value = "Total_Stock_Volume"

        ' Create summary variables
        Dim ticker_sym As String
        
        Dim yearly_change As Double
        
        Dim perc_change As Double
        
        Dim total_vol As Double
        total_vol = 0
        
        ' Track location for each ticker row
        Dim summary_table As Integer
        summary_table = 2
        
        ' Identify last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Create variables opening_price and closing_price
        Dim opening_price As Double
        Dim closing_price As Double
        
        ' Opening price for the very first ticker
        opening_price = ws.Cells(2, 3).Value
        
        ' For loop through all the stocks
        For i = 2 To LastRow
            ' Check if ticker is different from the next
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' ticker symbol
                ticker_sym = ws.Cells(i, 1).Value
                
                'very last value of closing price for ticker
                closing_price = ws.Cells(i, 6).Value
                
                ' yearly change
                yearly_change = (closing_price - opening_price)
                
                ' percentage change
                perc_change = (yearly_change / opening_price)
                
                'total stock volume
                total_vol = total_vol + ws.Cells(i, 7).Value
                
                ' Display values in Summary Table
                ws.Range("I" & summary_table).Value = ticker_sym
                
                ws.Range("J" & summary_table).Value = yearly_change
                ' If yearly change is negative - red. If positive - green
                If ws.Cells(summary_table, 10).Value < 0 Then
                    ws.Cells(summary_table, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(summary_table, 10).Interior.ColorIndex = 4
                End If
                
                ws.Range("K" & summary_table).Value = perc_change
                ws.Cells(summary_table, 11).NumberFormat = "0.00%"
                
                ws.Range("L" & summary_table).Value = total_vol
                
                ' Move to the next row of summary table
                summary_table = summary_table + 1
                
                ' Initialize next opening price and total volume
                ' Opening_price for next ticker
                opening_price = ws.Cells(i + 1, 3).Value
                
                ' total volume reset
                total_vol = 0
                
            ' If the same ticker symbol
            Else
                
                'total stock volume
                total_vol = total_vol + ws.Cells(i, 7).Value
            
            End If
            
        Next i
        
    Next ws
    
End Sub
