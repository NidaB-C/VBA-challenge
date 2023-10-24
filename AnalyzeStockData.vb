Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim startPrice As Double, endPrice As Double, yearlyChange As Double
    Dim percentChange As Double, totalVolume As Double
    Dim writeRow As Long
    Dim ticker As String
    Dim maxIncrease As Double, maxDecrease As Double, maxVolume As Double
    Dim maxIncreaseTicker As String, maxDecreaseTicker As String, maxVolumeTicker As String
    
    For Each ws In ThisWorkbook.Sheets
        ' Initialize variables
        writeRow = 2
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        
        ' Find last row in the sheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To lastRow
        
           ' Accumulate total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set end price and ticker
                endPrice = ws.Cells(i, 6).Value
                ticker = ws.Cells(i, 1).Value
                
                ' Calculate yearly change and percent change
                yearlyChange = endPrice - startPrice
                If startPrice <> 0 Then
                    percentChange = (yearlyChange / startPrice)
                Else
                    percentChange = 0
                End If
                
                ' Write data to columns
                ws.Cells(writeRow, 9).Value = ticker
                ws.Cells(writeRow, 10).Value = yearlyChange
                ws.Cells(writeRow, 11).Value = percentChange
                ws.Cells(writeRow, 11).NumberFormat = "0.00%"
                ws.Cells(writeRow, 12).Value = totalVolume
                
                ' Apply conditional formatting
                If yearlyChange > 0 Then
                    ws.Cells(writeRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(writeRow, 10).Interior.ColorIndex = 3
                End If

                ' Apply conditional formatting for Percent Change
                If percentChange > 0 Then
                ws.Cells(writeRow, 11).Interior.ColorIndex = 4
                Else
                ws.Cells(writeRow, 11).Interior.ColorIndex = 3
                End If
                
                ' Update maximum and minimum values
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                ElseIf percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If
                
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If
                
                ' Reset variables for next stock
                writeRow = writeRow + 1
                totalVolume = 0
                
                Else
                'Set start price for stock
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                    startPrice = ws.Cells(i, 3).Value
                End If
                
            End If
        Next i
        
        ' Output greatest increase, decrease, and volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 16).Value = maxIncreaseTicker
        ws.Cells(3, 16).Value = maxDecreaseTicker
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(2, 17).Value = maxIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = maxDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).Value = maxVolume
        
        ws.Columns("A:P").AutoFit
    Next ws
End Sub
