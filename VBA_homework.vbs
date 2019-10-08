Sub stock_value()

 'loop for each worksheet.
For Each WS In Worksheets

    'set heading to cloumns.
    WS.Cells(1, 11).Value = "Total Stock Volume"
    WS.Cells(1, 10).Value = "Ticker"
    WS.Cells(1, 12).Value = "Yearly Change"
    WS.Cells(1, 13).Value = "Percent change"
    WS.Cells(2, 14).Value = "Greatest % increase"
    WS.Cells(3, 14).Value = "Greatest % Decrease"
    WS.Cells(4, 14).Value = "Greatest total volume"
    
    'Initialize the variables.
    Dim total_vol, summaryrow As Integer
    Dim open_stock, close_stock, total_stock, year_percent, min, max, maxvol As Double
    close_stock = 0
    total_stock = 0
    min = 0
    maxvol = 0
    max = 0
    total_vol = WS.Cells(2, 7).Value
    open_stock = WS.Cells(2, 3).Value
    summaryrow = 2
    
    'Begin For loop that will iterate through all the rows in each sheet & calculate the required values.
    For I = 3 To 797712
        If (WS.Cells(I, 1).Value = WS.Cells(I - 1, 1).Value) Then
            total_vol = total_vol + WS.Cells(I, 7).Value
        
        ElseIf (WS.Cells(I, 1).Value <> WS.Cells(I - 1, 1).Value) Then
            close_stock = WS.Cells(I - 1, 6).Value
            total_stock = close_stock - open_stock
            
                'calculate year percent
                If (open_stock = 0) Then
                year_percent = 0
                Else
                year_percent = Round((total_stock / open_stock) * 100, 2)
                    
                    'calculate MIN & MAX values
                    If (year_percent < 0 And year_percent < min) Then
                        min = year_percent
                    ElseIf (year_percent > 0 And year_percent > max) Then
                        max = year_percent
                    End If
                End If
            open_stock = WS.Cells(I, 3).Value
            
            'calculate MAX VOLUME
            If (maxvol < total_vol) Then
                maxvol = total_vol
            End If
            
            'Entering values in Cloumns.
            WS.Cells(summaryrow, 11).Value = total_vol
            WS.Cells(summaryrow, 10).Value = Cells(I - 1, 1).Value
            WS.Cells(summaryrow, 12).Value = total_stock
                'Applying color formatting.
            If (total_stock < 0) Then
                WS.Cells(summaryrow, 12).Interior.Color = vbRed
            Else
                WS.Cells(summaryrow, 12).Interior.Color = vbGreen
            End If
            WS.Cells(summaryrow, 13).Value = year_percent & "%"
                summaryrow = summaryrow + 1
                total_vol = WS.Cells(I, 7).Value
                total_stock = 0
            
            WS.Cells(2, 15).Value = max
            WS.Cells(3, 15).Value = min
            WS.Cells(4, 15).Value = maxvol
        End If
    Next I
Next WS
End Sub