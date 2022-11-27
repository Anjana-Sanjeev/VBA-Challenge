Sub stock():

'To loop through each worksheets in the workbook at same time
For Each ws In Worksheets

    'Declaring variables to data types
    Dim lastrow, i, j, summary_table_row As Integer
    Dim ticker_name As String
    Dim total_stock As Double
    Dim yearly_change As Variant
    Dim yearly_change_start As Variant
    Dim percent_change As Variant
    
    'Finding last row of the data set
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Assigning headers for summary table
    ws.Cells(1, 11).Value = "Ticker"
    ws.Cells(1, 12).Value = "Yearly Change"
    ws.Cells(1, 13).Value = "Percent Change"
    ws.Cells(1, 14).Value = "Total Stock Volume"
    ws.Cells(2, 17).Value = "Greatest % Increase"
    ws.Cells(3, 17).Value = "Greatest % Decrease"
    ws.Cells(4, 17).Value = "Greatest Total Volume"
    ws.Cells(1, 18).Value = "Ticker"
    ws.Cells(1, 19).Value = "Value"
    
    'Formatting for summary table
    ws.Range("K1:N1").Font.Size = 12
    ws.Range("K1:N1").Font.Name = "Calibri"
    ws.Range("K1:Q1").EntireColumn.AutoFit
    
    'Assigning initial value for variables
    total_stock = 0
    yearly_change = 0
    percent_change = 0
    summary_table_row = 2
    
    
    
        'For looping through ticker, open value, close value and stock volume
        For i = 2 To lastrow
        
            'To find the first open value for calculating yearly change
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                yearly_change_start = ws.Cells(i, 3).Value
            
            End If
        
            'To find the last open value for calculating yearly change
            'To find percent change
            'To find total stock volume
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                ticker_name = ws.Cells(i, 1).Value
                yearly_change = ws.Cells(i, 6).Value - yearly_change_start
                percent_change = ((ws.Cells(i, 6).Value) / (yearly_change_start)) - 1
                total_stock = total_stock + ws.Cells(i, 7).Value
                
                ws.Range("K" & summary_table_row).Value = ticker_name
                ws.Range("L" & summary_table_row).Value = yearly_change
                ws.Range("M" & summary_table_row).Value = percent_change
                ws.Range("N" & summary_table_row).Value = total_stock
                
                summary_table_row = summary_table_row + 1
                
                'Resetting values to start loop again
                yearly_change = 0
                percent_change = 0
                total_stock = 0
                
            Else
               
                total_stock = total_stock + ws.Cells(i, 7).Value
                
            End If
        
        Next i
        
    
    
        'To assign conditional formatting for yealy change and percent change
        'Format yearly change into $ currency and percent change to %
        'To get max and min summary table
        For j = 2 To summary_table_row - 1
        
            ws.Cells(j, 13).Value = FormatPercent(ws.Cells(j, 13).Value, 2)
            
            
            
            If (ws.Cells(j, 12).Value) >= 0 Then
                    
            ws.Cells(j, 12).Interior.ColorIndex = 43
                            
            Else
                            
            ws.Cells(j, 12).Interior.ColorIndex = 3
                        
            End If
            
            
            
            If (ws.Cells(j, 13).Value) >= 0 Then
                    
            ws.Cells(j, 13).Interior.ColorIndex = 43
                            
            Else
                            
            ws.Cells(j, 13).Interior.ColorIndex = 3
                        
            End If
            
            
                
            ws.Cells(2, 19).Value = FormatPercent(Application.WorksheetFunction.Max(ws.Range("M2" & ":" & "M" & summary_table_row)), 2)
                    
            If ws.Cells(j, 13).Value = ws.Cells(2, 19).Value Then
                    
            ws.Cells(2, 18).Value = ws.Cells(j, 11).Value
                    
            End If
            
            
            
            
            If ws.Cells(j, 13).Value < 0 Then
                            
            ws.Cells(3, 19).Value = FormatPercent(Application.WorksheetFunction.Min(ws.Range("M2" & ":" & "M" & summary_table_row)), 2)
                           
            End If
                        
            If ws.Cells(j, 13).Value = ws.Cells(3, 19).Value Then
                        
            ws.Cells(3, 18).Value = ws.Cells(j, 11).Value
                        
            End If
            
            
            
            ws.Cells(4, 19).Value = Application.WorksheetFunction.Max(ws.Range("N2" & ":" & "N" & summary_table_row))
                            
            If ws.Cells(j, 14).Value = ws.Cells(4, 19).Value Then
                            
            ws.Cells(4, 18).Value = ws.Cells(j, 11).Value
                            
            End If
            
            
                            
            ws.Range("S4").EntireRow.AutoFit
            
        Next j
        
        ws.Range("L2" & ":" & "L" & summary_table_row).NumberFormat = "[$$-en-US]#,##0.00"

Next ws

End Sub

