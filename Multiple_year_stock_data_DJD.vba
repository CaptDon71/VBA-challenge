Sub StockMarketAnalysis()

    ' Declare variables
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim outputRow As Long
    Dim LastRow As Long
    Dim ticker As String
    Dim previousTicker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim firstRow As Long
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    
    ' Variables for tracking greatest values across all worksheets
    Dim maxPercentIncrease As Double
    Dim minPercentDecrease As Double
    Dim maxVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim minPercentDecreaseTicker As String
    Dim maxVolumeTicker As String

    ' Initialize tracking variables
    maxPercentIncrease = -1E+308
    minPercentDecrease = 1E+308
    maxVolume = 0
    maxPercentIncreaseTicker = ""
    minPercentDecreaseTicker = ""
    maxVolumeTicker = ""

    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Add headers for output in columns I to M if not already added
        If ws.Cells(1, 9).Value = "" Then
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
        End If

        ' Set the initial row for output
        outputRow = 2

        ' Determine the last row with data in column A (Ticker column)
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Initialize the row counter
        currentRow = 2

        ' Initialize variables
        previousTicker = ws.Cells(currentRow, 1).Value
        firstRow = currentRow
        openPrice = ws.Cells(currentRow, 3).Value
        totalVolume = 0

        ' Loop through all rows in the worksheet
        Do While currentRow <= LastRow
            ' Get the current ticker and check if it has changed
            ticker = ws.Cells(currentRow, 1).Value

            If ticker <> previousTicker Then
                ' Output the previous ticker's data
                closePrice = ws.Cells(currentRow - 1, 6).Value
                ' Calculate the quarterly change and percentage change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentageChange = (quarterlyChange / openPrice)
                Else
                    percentageChange = 0
                End If

                ' Output the results
                ws.Cells(outputRow, 9).Value = previousTicker
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
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 11).Value = percentageChange
                ws.Cells(outputRow, 12).Value = totalVolume

                ' Update the greatest values
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

                ' Move to the next output row
                outputRow = outputRow + 1

                ' Reset for the new ticker
                previousTicker = ticker
                openPrice = ws.Cells(currentRow, 3).Value
                totalVolume = 0
                firstRow = currentRow
            End If

            ' Accumulate the volume for the current ticker
            totalVolume = totalVolume + ws.Cells(currentRow, 7).Value

            ' Move to the next row
            currentRow = currentRow + 1
        Loop

        ' Calculate the quarterly change and percentage change for the last ticker
        closePrice = ws.Cells(currentRow - 1, 6).Value
        quarterlyChange = closePrice - openPrice
        If openPrice <> 0 Then
            percentageChange = (quarterlyChange / openPrice)
        Else
            percentageChange = 0
        End If

        ' Output the results for the last ticker
        ws.Cells(outputRow, 9).Value = previousTicker
        ws.Cells(outputRow, 10).Value = quarterlyChange
        ws.Cells(outputRow, 11).Value = percentageChange
        ws.Cells(outputRow, 12).Value = totalVolume

        ' Update the greatest values for the last ticker
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

    Next ws

    ' Output the greatest values only on the first worksheet
    With ThisWorkbook.Worksheets(1)
        .Cells(2, 15).Value = "Greatest % Increase"
        .Cells(3, 15).Value = "Greatest % Decrease"
        .Cells(4, 15).Value = "Greatest Total Volume"
        .Cells(1, 16).Value = "Ticker"
        .Cells(1, 17).Value = "Value"
        .Cells(2, 16).Value = maxPercentIncreaseTicker
        .Cells(3, 16).Value = minPercentDecreaseTicker
        .Cells(4, 16).Value = maxVolumeTicker
        .Cells(2, 17).Value = maxPercentIncrease
        .Cells(3, 17).Value = minPercentDecrease
        .Cells(4, 17).Value = maxVolume
    End With

    ' Display a message indicating that the job is complete
    MsgBox "Ticker stock data analysis is complete!"

End Sub

