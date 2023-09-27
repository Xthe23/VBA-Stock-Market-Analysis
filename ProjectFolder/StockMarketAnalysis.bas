Sub StockMarketAnalysis()

    ' Declare variables for worksheet, row numbers, and ticker information
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim ticker As String
    Dim yearOpen As Double, yearClose As Double
    Dim yearlyChange As Double, percentChange As Double
    Dim totalStockVolume As Double
    Dim startRow As Long

    ' Declare variables to keep track of greatest increase, decrease, and volume
    Dim greatestIncrease As Double, greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String, greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    ' Loop through each worksheet in the workbook
    For Each ws In Worksheets
        ' Assign headers for the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Initialize starting row and total volume
        startRow = 2
        totalStockVolume = 0
        yearOpen = ws.Cells(2, 3).Value  ' Initial opening price

        ' Initialize variables for tracking greatest values
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0

        ' Find the last row with data in the worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through each row in the worksheet
        For i = 2 To lastRow
            ' Add up the total stock volume
            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value

            ' Check for the last row of each ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = lastRow Then
                ' Assign ticker and yearClose values
                ticker = ws.Cells(i, 1).Value
                yearClose = ws.Cells(i, 6).Value

                ' Calculate yearly change and percent change
                yearlyChange = yearClose - yearOpen
                If yearOpen <> 0 Then
                    percentChange = yearlyChange / yearOpen
                Else
                    percentChange = 0
                End If

                ' Output the results in the summary table
                ws.Cells(startRow, 9).Value = ticker
                ws.Cells(startRow, 10).Value = yearlyChange
                ws.Cells(startRow, 11).Value = percentChange
                ws.Cells(startRow, 11).NumberFormat = "0.00%"
                ws.Cells(startRow, 12).Value = totalStockVolume

                ' Update greatest values and their corresponding tickers
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                ElseIf percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If
                If totalStockVolume > greatestVolume Then
                    greatestVolume = totalStockVolume
                    greatestVolumeTicker = ticker
                End If

                ' Apply conditional formatting for yearly change
                If yearlyChange > 0 Then
                    ws.Cells(startRow, 10).Interior.Color = RGB(0, 255, 0)  ' Green
                ElseIf yearlyChange < 0 Then
                    ws.Cells(startRow, 10).Interior.Color = RGB(255, 0, 0)  ' Red
                End If

                ' Prepare for the next ticker
                startRow = startRow + 1
                yearOpen = ws.Cells(i + 1, 3).Value
                totalStockVolume = 0
            End If
        Next i

        ' Output the greatest values in the additional summary table starting from column N
        ws.Cells(1, 14).Value = "Metrics"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = greatestIncreaseTicker
        ws.Cells(2, 16).Value = greatestIncrease
        ws.Cells(2, 16).NumberFormat = "0.00%"

        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = greatestDecreaseTicker
        ws.Cells(3, 16).Value = greatestDecrease
        ws.Cells(3, 16).NumberFormat = "0.00%"

        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = greatestVolumeTicker
        ws.Cells(4, 16).Value = greatestVolume

    Next ws  ' Move to the next worksheet in the workbook

End Sub

