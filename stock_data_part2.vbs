Attribute VB_Name = "Module2"
Sub findmaxmin()

'Bonus
    'loop in each worksheet
    For Each ws In Worksheets
        
        'define variables
        Dim r As Long
        Dim imax As Long
        Dim imin As Long
        Dim ivol As Long
        Dim lg As Double
        Dim sm As Double
        Dim lg_vol As Double
  
        'define the header of the table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
           
        'define the row of the table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
  
        'set the initial index of rows with max,min and max volumn as zero
        imax = 0
        imin = 0
        ivol = 0
        
        'find the last row of the summary table
        r = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'find the max value of greatest increase
        lg = Application.WorksheetFunction.Max(ws.Range("K2:K" & r))
        
        'find the min value of greatest decrease
        sm = Application.WorksheetFunction.Min(ws.Range("K2:K" & r))
        
        'find the value of largest volumn
        lg_vol = Application.WorksheetFunction.Max(ws.Range("L2:L" & r))
      
      'loop to find the rows number with greatest increase, greatest decrease and largest volume
      For i = 2 To r
        If ws.Cells(i, 11).Value = lg Then
           imax = i
        End If
            
        If ws.Cells(i, 11).Value = sm Then
            imin = i
        End If
        
        If ws.Cells(i, 12).Value = lg_vol Then
            ivol = i
        End If
      Next i
      
      'list the tickers and greatest increase, greatest decrease, and greates volume
        ws.Cells(2, 16).Value = ws.Cells(imax, 9).Value
        ws.Cells(2, 17).Value = FormatPercent(ws.Cells(imax, 11).Value)
        
        ws.Cells(3, 16).Value = ws.Cells(imin, 9).Value
        ws.Cells(3, 17).Value = FormatPercent(ws.Cells(imin, 11).Value)
        
        ws.Cells(4, 16).Value = ws.Cells(ivol, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(ivol, 12).Value
 Next ws

End Sub

