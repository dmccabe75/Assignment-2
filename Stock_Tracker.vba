Sub Stock_Tracker()

For Each ws In Worksheets

    Dim Ticker As String
    
    Dim Volume_Total As Double
    Volume_Total = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim Open_Price As Double
    Open_Price = ws.Cells(2, 3).Value
    
    Dim Close_Price As Double
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "% change"
    ws.Cells(1, 12).Value = "Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow

            'Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the ticker
            Ticker = ws.Cells(i, 1).Value
            
            ' Set Close_Price
            Close_Price = ws.Cells(i, 6).Value

             ' Add to the Volume Total
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value

            ' Print the Ticker in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker

            ' Print the Total Volume to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Volume_Total
            
            ' Print Yearly Change in Summary Table
            ws.Range("J" & Summary_Table_Row).Value = Close_Price - Open_Price
            
            ' Print Percentage Change in Summary Table
            ws.Range("K" & Summary_Table_Row).Value = FormatPercent((Close_Price - Open_Price) / Open_Price)

             ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            ' Reset the Volume Total
            Volume_Total = 0
            Open_Price = ws.Cells(1 + 1, 3).Value
            
            

            ' If the cell immediately following a row is the same brand...
            Else

             ' Add to the Volume Total
            Volume_Total = Volume_Total + Cells(i, 7).Value

            End If
            

    Next i
    
'Loop for Coloring
    
lastsummaryrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

    For j = 2 To lastsummaryrow
        If ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
        
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 4
            
        End If
        
    Next j
    
' Loop for Max %s and volume
    
Dim ticker_maxpct As String
Dim ticker_minpct As String
Dim ticker_maxvol As String

Dim maxpct As Double
maxpct = ws.Cells(2, 11).Value

Dim minpct As Double
minpct = ws.Cells(2, 11).Value

Dim maxvol As Double
maxvol = ws.Cells(2, 12).Value

lastsummaryrow = ws.Cells(Rows.Count, 11).End(xlUp).Row

    For k = 2 To lastsummaryrow
    
        If ws.Cells(k, 11).Value > maxpct Then
            maxpct = ws.Cells(k, 11).Value
            ticker_maxpct = ws.Cells(k, 9).Value
            
        End If
        
        If ws.Cells(k, 11).Value < minpct Then
            minpct = ws.Cells(k, 11).Value
            ticker_minpct = ws.Cells(k, 9).Value
            
        End If
    
        If ws.Cells(k, 12).Value > maxvol Then
            maxvol = ws.Cells(k, 12).Value
            ticker_maxvol = ws.Cells(k, 9).Value
            
        End If
        
    Next k
    
    'Set table values
    ws.Cells(2, 16).Value = ticker_maxpct
    ws.Cells(2, 17).Value = FormatPercent(maxpct)
    ws.Cells(3, 16).Value = ticker_minpct
    ws.Cells(3, 17).Value = FormatPercent(minpct)
    ws.Cells(4, 16).Value = ticker_maxvol
    ws.Cells(4, 17).Value = maxvol
    
    
Next ws

End Sub