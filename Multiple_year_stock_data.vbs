
Sub StockCounter()
    For Each ws In Worksheets


    Dim StockOpen As Double
    Dim StockVolume As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Ticker As String
    Dim NextTicker As String
    Dim column As Integer
    Dim currentOutRow As Integer
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    column = 1
    
    currentOutRow = 2
    
    StockVolume = 0
    StockOpen = 0
    
    
    ' Write Header Info
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Set Stock Open for 1st stock
    StockOpen = ws.Cells(2, 3).Value
    
    For i = 2 To lastrow
    
        Ticker = ws.Cells(i, 1).Value
        NextTicker = ws.Cells(i + 1, column).Value
        
        
       StockVolume = StockVolume + ws.Cells(i, 7).Value
        
        If Ticker <> NextTicker And Ticker <> "" Then
            ' Write the StockOpen and Percent Changes to Excel File
            If ws.Cells(i, 6).Value = 0 Or StockOpen = 0 Then
                PrecentChanged = 0
                YearlyChange = ws.Cells(i, 6).Value
              Else
                PercentChanged = (ws.Cells(i, 6).Value / StockOpen) - 1

                
                YearlyChange = ws.Cells(i, 6).Value - StockOpen
            End If
            
            
            ws.Cells(currentOutRow, 9).Value = Ticker
            ws.Cells(currentOutRow, 10).Value = YearlyChange
            
            If YearlyChange > 0 Then
                ws.Cells(currentOutRow, 10).Interior.ColorIndex = 4
            End If
            
            If YearlyChange < 0 Then
                ws.Cells(currentOutRow, 10).Interior.ColorIndex = 3
            End If
            
            ws.Cells(currentOutRow, 11).Value = PercentChanged
            ws.Cells(currentOutRow, 12).Value = StockVolume

            ' Set New Stock Numbers for next Ticker
            StockOpen = ws.Cells(i + 1, 3).Value
            StockVolume = 0
            currentOutRow = currentOutRow + 1
        
        
            End If
        
        
    Next i

    ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
    ws.Range("I2:I" & lastrow).NumberFormat = "0.00"


    Next ws


        
    
End Sub




