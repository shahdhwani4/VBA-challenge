Attribute VB_Name = "Module1"
Sub GenerateQuarterlyReportForWorksheets()
Dim ws As Worksheet

Set ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets

    ' Activate current ws
    ws.Activate
    
   ' Add required columns to table
    AddColumnsToTable
    
    ' Generate quaterly report
    GenerateQuarterlyChangeReport

  Next ws

End Sub

Sub GenerateQuarterlyChangeReport()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim currentTicker As String
    Dim startDate As Variant
    Dim endDate As Variant
    Dim startPrice As Double
    Dim endPrice As Double
    Dim quarterlyChange As Currency
    Dim outputRow As Long
    Dim quarterStartRow As Long
    Dim quarterStockVolume As Double
    Dim stockVolume As Double
    Dim gratestPrecentIncreaseValue As Double
    Dim gratestPrecentIncreaseTicker As String
    Dim gratestPercentDecreaseValue As Double
    Dim gratestPercentDecreaseTicker As String
    Dim gratestTotalVolumeValue As Double
    Dim gratestTotalVolumeTicker As String

    Set ws = ActiveSheet

    ' Find the last row of data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Initialize output row since headers are on row 1
    outputRow = 2
    
    ' Iterate through each row to find unique tickers and generate quarterly report
    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value
        
        ' If a new ticker is encountered or it's the first iteration
        If ticker <> currentTicker Or i = 2 Or Not IsInSameQuarter(startDate, ws.Cells(i, 2).Value) Then
            ' Update currentTicker
            currentTicker = ticker
            ' Reset quarterStartRow
            quarterStartRow = i
            ' Initialize quarterlyChange
            quarterlyChange = 0
            startDate = ws.Cells(i, 2).Value
            startPrice = ws.Cells(quarterStartRow, 3).Value
            quarterStockVolume = 0
        End If
        
        ' Get endDate, endPrice and stockVolume for current date
        endDate = ws.Cells(i, 2).Value
        endPrice = ws.Cells(i, 6).Value
        stockVolume = ws.Cells(i, 7).Value
        
        ' Compute quarter stock volume
        quarterStockVolume = quarterStockVolume + ws.Cells(i, 7).Value
        
        ' Caclculate quarterly change
        If IsNumeric(startPrice) And IsNumeric(endPrice) Then
            quarterlyChange = endPrice - startPrice
            percentChange = (quarterlyChange / startPrice) * 100
        Else
            quarterlyChange = 0
        End If
        
        ' If next row is not in same quarter or it's the last row of the current ticker or the last row of the data
        If Not IsInSameQuarter(startDate, ws.Cells(i + 1, 2).Value) Or i = lastRow Or ws.Cells(i + 1, 1).Value <> ticker Then
            ' Write quarter result to ws
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = Format(quarterlyChange, "0.00")
            ws.Cells(outputRow, 11).Value = Format(percentChange, "0.00") & "%"
            ws.Cells(outputRow, 12).Value = quarterStockVolume
            
            ' Apply formatting (Color) to quarterly change column
            If quarterlyChange > 0 Then
                ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
            ElseIf quarterlyChange < 0 Then
                ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
            Else
                ws.Cells(outputRow, 11).Interior.ColorIndex = xlNone ' no formattting
            End If
            
            ' Update gratestPrecentIncreaseValue
            If percentChange > gratestPrecentIncreaseValue Then
              gratestPrecentIncreaseValue = percentChange
              gratestPrecentIncreaseTicker = ticker
            End If
            
            ' Update gratestPercentDecreaseValue
            If percentChange < gratestPercentDecreaseValue Then
              gratestPercentDecreaseValue = percentChange
              gratestPercentDecreaseTicker = ticker
            End If
            
            ' Update gratestTotalVolumeValue
            If quarterStockVolume > gratestTotalVolumeValue Then
              gratestTotalVolumeValue = quarterStockVolume
              gratestTotalVolumeTicker = ticker
            End If
            
            outputRow = outputRow + 1
        End If
    Next i
    
    ' Write ticker stats to ws
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Increase"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(2, 16).Value = gratestPrecentIncreaseTicker
    ws.Cells(3, 16).Value = gratestPercentDecreaseTicker
    ws.Cells(4, 16).Value = gratestTotalVolumeTicker
    ws.Cells(2, 17).Value = Format(gratestPrecentIncreaseValue, "0.00") & "%"
    ws.Cells(3, 17).Value = Format(gratestPercentDecreaseValue, "0.00") & "%"
    ws.Cells(4, 17).Value = Format(gratestTotalVolumeValue, "0.00")
    
    ' Autofit ws columns
    ws.Columns(9).AutoFit
    ws.Columns(10).AutoFit
    ws.Columns(11).AutoFit
    ws.Columns(12).AutoFit
    ws.Columns(15).AutoFit
    ws.Columns(16).AutoFit
    ws.Columns(17).AutoFit
End Sub


Function IsInSameQuarter(startDate As Variant, currentDate As Variant) As Boolean
    Dim startQuarter As Integer
    Dim currentQuarter As Integer
    Dim startYear As Integer
    Dim currentYear As Integer
    Dim startMonth As Integer
    Dim currentMonth As Integer
    
    ' Convert startDate and currentDate  to Date type
    startDate = ParseDate(startDate)
    currentDate = ParseDate(currentDate)

    ' Get year and month from start date
    startYear = year(startDate)
    startMonth = month(startDate)
    ' Compute quarter based on the month
    startQuarter = Int((startMonth - 1) / 3) + 1
    
    ' Get year and month from current date
    currentYear = year(currentDate)
    currentMonth = month(currentDate)
    ' Determine the quarter based on the month
    currentQuarter = Int((currentMonth - 1) / 3) + 1

    ' Compare quarters and years
    IsInSameQuarter = (startQuarter = currentQuarter) And (startYear = currentYear)
End Function

Function ParseDate(dateValue As Variant) As Date
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer

    ' Check if the date is in "20200102" format
    If IsNumeric(dateValue) And Len(dateValue) = 8 Then
        year = Left(dateValue, 4)
        month = Mid(dateValue, 5, 2)
        day = Right(dateValue, 2)
        ParseDate = DateSerial(year, month, day)
    ElseIf IsDate(dateValue) Then
        ' In date format like "1/2/22"
        ParseDate = CDate(dateValue)
    Else
        ' If the date format is not recognized, do nothing
        ParseDate = dateValue
    End If
End Function

Sub AddColumnsToTable()
    
Dim ws As Worksheet

Set ws = ActiveSheet

lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
numColumns = 10

For i = 1 To numColumns
  ws.Columns(lastCol + i).Insert
  Next i

' Add column label for column at row 1
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

End Sub



