Sub StockMarketAnalysis()
    ' Declare variables
    Dim ws As Worksheet
    Dim currentRow As Long, outputRow As Long, LastRow As Long
    Dim ticker As String, previousTicker As String
    Dim openPrice As Double, closePrice As Double, totalVolume As Double
    Dim firstRow As Long, quarterlyChange As Double, percentageChange As Double
    
    ' Variables for tracking greatest values per worksheet
    Dim maxPercentIncrease As Double, minPercentDecrease As Double, maxVolume As Double
    Dim maxPercentIncreaseTicker As String, minPercentDecreaseTicker As String, maxVolumeTicker As String

    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Reset greatest values for each worksheet (each quarter)
        maxPercentIncrease = -1E+308
        minPercentDecrease = 1E+308
        maxVolume = 0
        maxPercentIncreaseTicker = ""
        minPercentDecreaseTicker = ""
        maxVolumeTicker = ""

        ' Add headers if not already added
        If ws.Cells(1, 9).Value = "" Then
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
        End If

        outputRow = 2  ' Initial output row
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        currentRow = 2  ' Start from second row

        previousTicker = ws.Cells(currentRow, 1).Value
        firstRow = currentRow
        openPrice = ws.Cells(currentRow, 3).Value
        totalVolume = 0

        ' Loop through rows
        Do While currentRow <= LastRow
            ticker = ws.Cells(currentRow, 1).Value

            If ticker <> previousTicker Then
                ' Calculate previous ticker's values
                closePrice = ws.Cells(currentRow - 1, 6).Value
                quarterlyChange = closePrice - openPrice
                percentageChange = IIf(openPrice <> 0, (quarterlyChange / openPrice) * 100, 0)

                ' Output the results
                ws.Cells(outputRow, 9).Value = previousTicker
                ws.Cells(outputRow, 10).Value = Round(quarterlyChange, 2)
                ws.Cells(outputRow, 11).Value = Round(percentageChange, 2)
                ws.Cells(outputRow, 12).Value = totalVolume

                ' Conditional formatting
                If quarterlyChange < 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3 ' Red
                    ws.Cells(outputRow, 11).Interior.ColorIndex = 3 ' Red
                ElseIf quarterlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4 ' Green
                    ws.Cells(outputRow, 11).Interior.ColorIndex = 4 ' Green
                Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 2 ' White
                    ws.Cells(outputRow, 11).Interior.ColorIndex = 2 ' White
                End If

                ' Update greatest values for this worksheet
                If percentageChange > maxPercentIncrease Then
                    maxPercentIncrease = percentageChange
                    maxPercentIncreaseTicker = previousTicker
                End If
                If percentageChange < minPercentDecrease Then
                    minPercentDecrease = percentageChange
                    minPercentDecreaseTicker = previousTicker
                End If
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = previousTicker
                End If

                ' Prepare for new ticker
                outputRow = outputRow + 1
                previousTicker = ticker
                openPrice = ws.Cells(currentRow, 3).Value
                totalVolume = 0
                firstRow = currentRow
            End If

            ' Accumulate volume
            totalVolume = totalVolume + ws.Cells(currentRow, 7).Value
            currentRow = currentRow + 1
        Loop

        ' Handle last ticker in the sheet
        closePrice = ws.Cells(currentRow - 1, 6).Value
        quarterlyChange = closePrice - openPrice
        percentageChange = IIf(openPrice <> 0, (quarterlyChange / openPrice) * 100, 0)

        ' Output last ticker's results
        ws.Cells(outputRow, 9).Value = previousTicker
        ws.Cells(outputRow, 10).Value = Round(quarterlyChange, 2)
        ws.Cells(outputRow, 11).Value = Round(percentageChange, 2)
        ws.Cells(outputRow, 12).Value = totalVolume

        ' Update greatest values for the last ticker
        If percentageChange > maxPercentIncrease Then
            maxPercentIncrease = percentageChange
            maxPercentIncreaseTicker = previousTicker
        End If
        If percentageChange < minPercentDecrease Then
            minPercentDecrease = percentageChange
            minPercentDecreaseTicker = previousTicker
        End If
        If totalVolume > maxVolume Then
            maxVolume = totalVolume
            maxVolumeTicker = previousTicker
        End If

        ' Output the greatest values for this quarter in the worksheet
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 16).Value = maxPercentIncreaseTicker
        ws.Cells(3, 16).Value = minPercentDecreaseTicker
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(2, 17).Value = Round(maxPercentIncrease, 2)
        ws.Cells(3, 17).Value = Round(minPercentDecrease, 2)
        ws.Cells(4, 17).Value = maxVolume

    Next ws

    ' Final completion message
    MsgBox "Stock market analysis is complete for all quarters!"

End Sub
