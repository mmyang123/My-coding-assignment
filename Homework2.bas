Attribute VB_Name = "Module1"
Sub StocksTicker()
    Dim currentTicker As String     ' Current stock ticker
    Dim nextTicker As String        ' Next ticker from next row
    Dim ticker As String            ' Current row's stock ticker
    Dim openingPrice As Double      ' Opening price from first day of year
    Dim closingPrice As Double      ' Closing price from last day of year
    Dim yearlyChange As Double      ' Equals closingPrice - openingPrie
    Dim percentChange As Double     ' Equals yearlyChange/openingPrice
    Dim totalStockVolume As Variant    ' Sum of volume for year
    
    Dim lastRow As Long             ' Last row in sheet
    Dim outputRow As Integer        ' Row to write the output to
    
    For Each ws In Worksheets
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        currentTicker = ""
        outputRow = 2
        
        ' Generate labels for output data
        ws.Cells(1, 9).Value = "ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For rw = 2 To lastRow
            ticker = ws.Cells(rw, 1).Value
            If (ticker <> currentTicker) Then
                currentTicker = ticker
                openingPrice = ws.Cells(rw, 3).Value
                totalStockVolume = ws.Cells(rw, 7).Value
                ws.Cells(outputRow, 9).Value = ticker
            Else
                totalStockVolume = totalStockVolume + ws.Cells(rw, 7).Value
            End If
            
            ' Find if we are at the last row for the current ticker
            If (rw = lastRow) Then
                closingPrice = ws.Cells(rw, 6).Value
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice
                ws.Cells(outputRow, 10).Value = yearlyChange
                If (yearlyChange < 0) Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                End If
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalStockVolume
                outputRow = outputRow + 1
            Else
                If (rw < lastRow) Then
                    nextTicker = ws.Cells(rw + 1, 1).Value
                    If (ticker <> nextTicker) Then
                        closingPrice = ws.Cells(rw, 6).Value
                        yearlyChange = closingPrice - openingPrice
                        percentChange = yearlyChange / openingPrice
                        ws.Cells(outputRow, 10).Value = yearlyChange
                        If (yearlyChange < 0) Then
                            ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                        Else
                            ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                        End If
                        ws.Cells(outputRow, 11).Value = percentChange
                        ws.Cells(outputRow, 12).Value = totalStockVolume
                        outputRow = outputRow + 1
                    End If
                End If
            End If
        Next rw
    Next ws
End Sub

