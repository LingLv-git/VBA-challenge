Attribute VB_Name = "Module1"

Sub stock_summary()
    'loop in worksheets
    For Each ws In Worksheets
        'define variables
        Dim vol_total As Variant
        Dim lRow As Long
        Dim lCol As Long
        Dim i As Long
        Dim openprice As Double
        Dim closeprice As Double
    
        Dim summary_table_row As Integer
        summary_table_row = 2
        
        Dim ticker As String
   
        vol_total = 0
        
        'find the open price for each worksheet
        openprice = 0
        
        
        'finds the last row
        lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
        'finds the last column
        lCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column

    
        'add header to each worksheet
        ws.Cells(1, lCol + 2).Value = "Ticker"
        ws.Cells(1, lCol + 3).Value = "Yearly Change"
        ws.Cells(1, lCol + 4).Value = "Percent Change"
        ws.Cells(1, lCol + 5).Value = "Total Stock Volumn"
                  
        'loop from second row to the last row
        For i = 2 To lRow
            
            'check if open price is 0, reset openprice to the ith value of the cell
            If openprice = 0 Then
                openprice = ws.Cells(i, 3).Value
            End If
            
            'check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               
                ticker = ws.Cells(i, 1).Value
                                    
                vol_total = vol_total + ws.Cells(i, 7).Value
                
                'assign ticker to a summary table
                ws.Cells(summary_table_row, lCol + 2).Value = ticker
                
                'assign price yearly change
                ws.Cells(summary_table_row, lCol + 3).Value = FormatNumber(ws.Cells(i, 6).Value - openprice, 2)
                
                'fill the cell background color: positive is green(4) and negative is red(3)
                If ws.Cells(summary_table_row, lCol + 3).Value >= 0 Then
                    ws.Cells(summary_table_row, lCol + 3).Interior.ColorIndex = 4
                 Else
                    ws.Cells(summary_table_row, lCol + 3).Interior.ColorIndex = 3
                End If
                    
                If openprice = 0 Then
                    ws.Cells(summary_table_row, lCol + 4).Value = ""
                Else
                    ws.Cells(summary_table_row, lCol + 4).Value = FormatPercent(ws.Cells(summary_table_row, lCol + 3).Value / openprice, 2)
                End If
            
                ws.Cells(summary_table_row, lCol + 5).Value = vol_total
                
                summary_table_row = summary_table_row + 1
                
                'reset the vol_total
                vol_total = 0
                
                'reset the open price
                openprice = 0

                Else
                    vol_total = vol_total + ws.Cells(i, 7).Value
             
             End If
         
         Next i
         
   Next ws

End Sub


