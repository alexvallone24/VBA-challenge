Attribute VB_Name = "Module1"
Sub EnhancedStockAnalysis()
    ' Declare our variables - think of this as setting up our toolbox
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim ticker As String
    Dim openingPrice As Double, closingPrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    
    ' Set up our worksheet and find last row
    For Each ws In ThisWorkbook.Worksheets
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Create our summary headers - this is like creating column titles
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    ' Initialize our tracking variables
    Dim currentRow As Long
    currentRow = 2  ' Start at row 2 since row 1 has headers
    totalVolume = 0
    openingPrice = ws.Cells(2, 3).Value  ' Get first opening price
    ' Loop through all rows of data
    For i = 2 To lastRow
        ' Check if we're still on the same ticker
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            ' Add to volume but keep tracking same ticker
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        Else
            ' We've reached the end of current ticker's data
            ' Add final volume for this ticker
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Get closing price for this ticker
            closingPrice = ws.Cells(i, 6).Value
            
            ' Calculate our changes
            quarterlyChange = closingPrice - openingPrice
            
            ' Calculate percent change with error handling
            If openingPrice <> 0 Then
                percentChange = quarterlyChange / openingPrice
            Else
                percentChange = 0
            End If
            ' Output all our calculations
            ws.Cells(currentRow, 9).Value = ws.Cells(i, 1).Value  ' Ticker
            ws.Cells(currentRow, 10).Value = quarterlyChange      ' $ Change
            ws.Cells(currentRow, 11).Value = percentChange        ' % Change
            ws.Cells(currentRow, 12).Value = totalVolume          ' Volume
            ' Format percent as percentage
            ws.Cells(currentRow, 11).NumberFormat = "0.00%"
            
            ' Add color formatting - Green for positive, Red for negative
            If quarterlyChange > 0 Then
                ws.Cells(currentRow, 10).Interior.ColorIndex = 4  ' Green
            ElseIf quarterlyChange < 0 Then
                ws.Cells(currentRow, 10).Interior.ColorIndex = 3  ' Red
            End If
            
            ' Reset for next ticker
            currentRow = currentRow + 1
            totalVolume = 0
            openingPrice = ws.Cells(i + 1, 3).Value  ' Get next ticker's opening price
        End If
    Next i
    ' Create summary section
    ws.Cells(1, 15).Value = "Summary Statistics"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ' Calculate summary statistics
    Dim maxIncrease As Range, maxDecrease As Range, maxVolume As Range
    Set maxIncrease = ws.Range("K2:K" & currentRow - 1)
    Set maxDecrease = ws.Range("K2:K" & currentRow - 1)
    Set maxVolume = ws.Range("L2:L" & currentRow - 1)
    
    ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(maxIncrease)
    ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(maxDecrease)
    ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(maxVolume)
    
    ' Format summary percentages
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).NumberFormat = "0.00%"
    ' Add corresponding ticker symbols
    Dim maxRow As Long, minRow As Long, volRow As Long
    maxRow = Application.Match(ws.Cells(2, 16).Value, maxIncrease, 0) + 1
    minRow = Application.Match(ws.Cells(3, 16).Value, maxDecrease, 0) + 1
    volRow = Application.Match(ws.Cells(4, 16).Value, maxVolume, 0) + 1
    
    ws.Cells(2, 15).Value = ws.Cells(maxRow, 9).Value
    ws.Cells(3, 15).Value = ws.Cells(minRow, 9).Value
    ws.Cells(4, 15).Value = ws.Cells(volRow, 9).Value
    
    Next ws
    
End Sub

