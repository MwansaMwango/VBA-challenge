Sub analyseStocks()
    Dim ws As Worksheet

    For Each ws In Worksheets
    
   
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        nextRowTicker = ws.Cells(Rows.Count, "I").End(xlUp).Row + 1
        totalStockVolume = 0
        greatestPercentageIncrease = 0
        greatestPercentageDecrease = 0
        greatestStockVolume = 0
        
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim yearPriceChangePercentage As Double
        openingPrice = False 'checks for new stock opening price - assumes price change comparison from opening price at start of year
        
        With ws
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percentage Change"
            .Range("L1").Value = "Total Stock Volume"
            .Range("N1").Value = "Metric"
            .Range("N2").Value = "Greatest Increase"
            .Range("N3").Value = "Greatest Decrease"
            .Range("N4").Value = "Greateast Total Volume"
            .Range("O1").Value = "Ticker"
            .Range("P1").Value = "Value"
               
            For i = 2 To lastRow
            
                If openingPrice = False Then
                     openingPrice = .Range("C" & i).Value
                End If
        
                
                If .Range("A" & i).Value = .Range("A" & i + 1).Value Then
                  
                    totalStockVolume = totalStockVolume + .Range("G" & i).Value
                
                   
                Else
                    
                    closingPrice = Range("F" & i).Value
                   
                    yearPriceChange = closingPrice - openingPrice
                    If openingPrice > 0 Then
                    
                    yearPriceChangePercentage = (yearPriceChange / openingPrice)
                    End If
                    
                    'Calculate greatest stock stats
                    If yearPriceChangePercentage > greatestPercentageIncrease Then
                        greatestPercentageIncrease = yearPriceChangePercentage
                        greatestPercentageIncreaseTicker = .Range("A" & i).Value
                    End If
                    
                    If yearPriceChangePercentage < greatestPercentageDecrease Then
                        greatestPercentageDecrease = yearPriceChangePercentage
                        greatestPercentageDecreaseTicker = .Range("A" & i).Value
                    End If
                    
                    totalStockVolume = totalStockVolume + .Range("G" & i).Value
                    If totalStockVolume > greatestStockVolume Then
                        greatestStockVolume = totalStockVolume
                        greatestStockVolumeTicker = .Range("A" & i).Value
                    End If
                    
                    'Print stock stats
                    .Range("i" & nextRowTicker).Value = .Range("A" & i).Value
                    .Range("j" & nextRowTicker).Value = yearPriceChange
                    .Range("k" & nextRowTicker).Value = yearPriceChangePercentage
                    .Range("l" & nextRowTicker).Value = totalStockVolume
                    
                    'Conditioinal formating
                    If yearPriceChangePercentage < 0 Then
                        .Range("k" & nextRowTicker).Interior.ColorIndex = 3 'red
                    Else
                        .Range("k" & nextRowTicker).Interior.ColorIndex = 4 'green
                    End If
                    
                    'Next row
                    nextRowTicker = nextRowTicker + 1
                    
                    'Zero stats
                    totalStockVolume = 0
                    openingPrice = False 'use new opening price for new stock
                
                End If
                
            Next i
        
            'Print Greatest Tickers
            .Range("o2").Value = greatestPercentageIncreaseTicker
            .Range("o3").Value = greatestPercentageDecreaseTicker
            .Range("o4").Value = greatestStockVolumeTicker
            
            'Print Greatest Values
            .Range("p2").Value = greatestPercentageIncrease
            .Range("p3").Value = greatestPercentageDecrease
            .Range("p2:p3").NumberFormat = "0.00%"
            .Range("p4").Value = greatestStockVolume
            .Range("p4").NumberFormat = "0,000"
            
            'Conditional formating
            If greatestPercentageIncrease < 0 Then
                .Range("p2").Interior.ColorIndex = 3 'red
            Else
                .Range("p2").Interior.ColorIndex = 4 'green
            End If
            
            If greatestPercentageDecrease < 0 Then
                .Range("p3").Interior.ColorIndex = 3 'red
            Else
                .Range("p3").Interior.ColorIndex = 4 'green
            End If
            ' Data presentation formating
            .Columns("A:P").AutoFit
            .Range("K2:K" & lastRow).NumberFormat = "0.00%"
            .Range("L2:L" & lastRow).NumberFormat = "0,000"
            
        End With
    Next ws

    
End Sub

