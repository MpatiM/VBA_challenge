# VBA_challenge
The Module 2 assignment was completed to assess the skills and understanding of VBA.

## First Task
Create a script that loops through all the stocks for one year and outputs the following information:
* The ticker symbol
* Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
* The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
* The total stock volume of the stock.

The code:
```
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
```
## Second Task
Add functionality to your script to return the stock with the 
* "Greatest % increase"
* "Greatest % decrease"
* "Greatest total volume"

The code:
```
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
```
## Results
The results given are shown in the following images for each year
* 2018
  ![stocks_summary_2018](https://github.com/MpatiM/VBA_challenge/assets/159741444/7e59e7ca-7e09-41b8-b496-b54980784ecd)
* 2019
![stocks_summary_2019](https://github.com/MpatiM/VBA_challenge/assets/159741444/76886220-0549-41be-be24-12d275b8a379)
* 2020
![stocks_summary_2020](https://github.com/MpatiM/VBA_challenge/assets/159741444/0b4aec4f-7a31-4b04-a120-cdecdc745064)


