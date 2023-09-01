# VBA-challenge
Sub Module2_Challenge()

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim LastRowA As Double
        Dim LastRowB As Double
        Dim NewRow As Integer
        
        Dim TickerSymbol As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim StockVolume As Double
        Dim GreatIncrease As Double
        Dim GreatDecrease As Double
        Dim TotalVolume As Double
        
        WorksheetName = ws.Name
        
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Symbol"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        TickerSymbol = 2
        NewRow = 2
        
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
        For i = 2 To LastRowA
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickerSymbol, 9).Value = ws.Cells(i, 1).Value
                OpenPrice = ws.Cells(i, 3).Value
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                PercentChange = (YearlyChange / OpenPrice) * 100
                StockVolume = WorksheetFunction.Sum(ws.Range(ws.Cells(i, 7), ws.Cells(NewRow, 7)))
                 
                ws.Cells(TickerSymbol, 10).Value = YearlyChange
                ws.Cells(TickerSymbol, 11).Value = PercentChange
                ws.Cells(TickerSymbol, 12).Value = StockVolume
               
                If ws.Cells(TickerSymbol, 10).Value < 0 Then
                    ws.Cells(TickerSymbol, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(TickerSymbol, 10).Interior.ColorIndex = 4
                End If
            
            TickerSymbol = TickerSymbol + 1
            NewRow = NewRow + 1
            
            End If
                   
        Next i
    
    LastRowB = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRowB
        If ws.Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncrease = ws.Cells(i, 11)
            ws.Cells(2, 16).Value = GreatestIncrease
        Else
            GreatIncrease = GreatIncrease
                
        End If
            
        If ws.Cells(i, 11).Value < GreatestDecreaase Then
            GreatestDecrease = ws.Cells(i, 11)
            ws.Cells(3, 16).Value = GreatestDecrease
        Else
            GreatestDecrease = GreatestDecrease
            
        End If
            
        If ws.Cells(i, 12).Value > TotalVolume Then
            TotalVolume = ws.Cells(i, 12).Value
            ws.Cells(4, 17).Value = ws.Cells(i, 9).Value
        Else
            TotalVolume = TotalVolume
                
        End If
        
    Next i
        
    Next ws

End Sub
