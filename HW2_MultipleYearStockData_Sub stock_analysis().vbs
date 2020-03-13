Sub stock_analysis()
    
Set starting_ws = ActiveSheet

Dim yearly_change, percent_change, opening_price, closing_price As Double
Dim lastrow, total_stock_volume As Double
Dim summary_stock_table_row As Integer
        

For Each ws In Worksheets

    total_stock_volume = 0
    summary_stock_table_row = 2
    ws.Range("J1") = "Stock"
    ws.Range("K1") = "Yearly_Change"
    ws.Range("L1") = "Percent_Change"
    ws.Range("M1") = "Total_Stock_Volume"
    ws.Range("Q1") = "Stock"
    ws.Range("R1") = "Value"
    ws.Range("P2") = "Greatest % Increase"
    ws.Range("P3") = "Greatest % Decrease"
    ws.Range("P4") = "Greatest Total Volume"
                
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To lastrow
        If (ws.Cells(i + 1, 1) <> ws.Cells(i, 1)) Then
            
            ws.Range("J" & summary_stock_table_row).Value = ws.Cells(i, 1).Value
        
            closing_price = ws.Cells(i, 6).Value
                               
            yearly_change = closing_price - opening_price
            
            'Color formatting cells based on yearly change
            If yearly_change < 0 Then
            
                ws.Range("K" & summary_stock_table_row).Value = yearly_change
                ws.Range("K" & summary_stock_table_row).Interior.Color = 255
            Else
                ws.Range("K" & summary_stock_table_row).Value = yearly_change
                ws.Range("K" & summary_stock_table_row).Interior.Color = 5296274
            
            End If
                        
            'In case opening price equals zero
            If yearly_change = closing_price Then
            
                percent_change = FormatPercent(yearly_change, 2)
            
            Else
            
                percent_change = FormatPercent((closing_price - opening_price) / opening_price, 2)
            
            End If
            
            ws.Range("L" & summary_stock_table_row).Value = percent_change
            
            ws.Range("M" & summary_stock_table_row).Value = total_stock_volume + ws.Cells(i, 7)
                
            total_stock_volume = 0
            
            summary_stock_table_row = summary_stock_table_row + 1
                        
        ElseIf (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) And (ws.Cells(i, 2).Value = "20160101") Then
                
            opening_price = ws.Cells(i, 3).Value
            
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
        ElseIf (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) And (ws.Cells(i, 2).Value = "20150101") Then
                
            opening_price = ws.Cells(i, 3).Value
            
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
        ElseIf (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) And (ws.Cells(i, 2).Value = "20140101") Then
                
            opening_price = ws.Cells(i, 3).Value
            
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
        Else
        
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
        End If
            
                    
    Next i
    
    ' Determine greatest percent increase, decrease and total volume
    lastrow_percent = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    
    max_percent_increase = 0
    min_percent_decrease = 0
    max_greatest_volume = 0
    
    
    'Determine greatest percent increase
    For i = 2 To lastrow_percent
        
        If ws.Cells(i, "L") > max_percent_increase Then
            
            max_percent_increase = ws.Cells(i, "L")
        
        Else
            
            max_percent_increase = max_percent_increase
            
        End If
    
    Next i
            
    'Determine greatest precent decrease
    For i = 2 To lastrow_percent
    
        If ws.Cells(i, "L") < min_percent_decrease Then
        
            min_percent_decrease = ws.Cells(i, "L")
        
        Else
        
            min_percent_decrease = min_percent_decrease
        
        End If
    
    Next i
    
    'Determine greatest total volume
    For i = 2 To lastrow_percent
    
        If ws.Cells(i, "L").Offset(0, 1).Value > max_greatest_volume Then
        
            max_greatest_volume = ws.Cells(i, "L").Offset(0, 1).Value
        
        Else
        
            max_greatest_volume = max_greatest_volume
            
        End If
        
    Next i
    
ws.Range("R2").Value = FormatPercent(max_percent_increase, 2)
ws.Range("R3").Value = FormatPercent(min_percent_decrease, 2)
ws.Range("R4").Value = max_greatest_volume

    'Determine stock corresponding to greatest percent increase, decrease and total volume
    
    For i = 2 To lastrow_percent

        If ws.Cells(i, "L").Value = ws.Range("R2").Value Then

            ws.Range("Q2").Value = ws.Cells(i, "L").Offset(0, -2).Value

        ElseIf ws.Cells(i, "L").Value = ws.Range("R3").Value Then

            ws.Range("Q3").Value = ws.Cells(i, "L").Offset(0, -2).Value

        ElseIf ws.Cells(i, "L").Offset(0, 1).Value = ws.Range("R4").Value Then

            ws.Range("Q4").Value = ws.Cells(i, "L").Offset(0, -2).Value

        End If

    Next i

                
Next ws
   
starting_ws.Activate

End Sub
