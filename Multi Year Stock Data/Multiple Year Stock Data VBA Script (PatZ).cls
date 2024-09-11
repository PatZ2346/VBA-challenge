Sub QuarterlyStockAnalysisAllSheets()
    Dim ws As Worksheet, lastRow As Long, i As Long, startRow As Long, endRow As Long, outputRow As Long
    Dim ticker As String, quarter As String
    Dim openPrice As Double, closePrice As Double, totalVolume As Double, change As Double, percentChange As Double
    Dim maxIncreaseTicker As String, maxDecreaseTicker As String, maxVolumeTicker As String
    Dim maxIncreaseValue As Double, maxDecreaseValue As Double, maxVolumeValue As Double
    
    ' My first loop to process each worksheet and calculate values
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        outputRow = 2
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            quarter = Format(ws.Cells(i, 2).Value, "yyyy\Qq")
            
            startRow = i
            Do While ws.Cells(i, 1).Value = ticker And Format(ws.Cells(i, 2).Value, "yyyy\Qq") = quarter
                i = i + 1
            Loop
            endRow = i - 1
            i = i - 1
            
            openPrice = ws.Cells(startRow, 3).Value
            closePrice = ws.Cells(endRow, 6).Value
            totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))
            change = closePrice - openPrice
            percentChange = (change / openPrice) * 100
            
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = Format(change, "0")
            ws.Cells(outputRow, 11).Value = Format(percentChange, "0.00") & "%"
            ws.Cells(outputRow, 12).Value = Format(totalVolume, "0")
            
            ws.Cells(outputRow, 10).Interior.Color = IIf(change > 0, RGB(0, 255, 0), RGB(255, 0, 0))
            
            outputRow = outputRow + 1
        Next i
        
        ' Initialize the greatest values for each worksheet
        maxIncreaseValue = -1
        maxDecreaseValue = 1
        maxVolumeValue = 0
        
        ' My second loop to find the greatest values within each worksheet
        For i = 2 To lastRow
            percentChange = ws.Cells(i, 11).Value
            totalVolume = ws.Cells(i, 12).Value
            ticker = ws.Cells(i, 9).Value
            
            If percentChange > maxIncreaseValue Then
                maxIncreaseValue = percentChange
                maxIncreaseTicker = ticker
            End If
            
            If percentChange < maxDecreaseValue Then
                maxDecreaseValue = percentChange
                maxDecreaseTicker = ticker
            End If
            
            If totalVolume > maxVolumeValue Then
                maxVolumeValue = totalVolume
                maxVolumeTicker = ticker
            End If
        Next i
        
        ' Then output the greatest values for each worksheet
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = maxIncreaseTicker
        ws.Cells(2, 17).Value = maxIncreaseValue
        ws.Cells(2, 17).NumberFormat = "0.00%;-0.00%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = maxDecreaseTicker
        ws.Cells(3, 17).Value = maxDecreaseValue
        ws.Cells(3, 17).NumberFormat = "0.00%;-0.00%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(4, 17).Value = Format(maxVolumeValue, "0")
    Next ws
End Sub


