Sub CreateSummaryTable()
    Dim objWorkbook As Workbook
    Dim objSheet As Worksheet
    Dim objSummarySheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim grandTotal As Double, totalDelayed As Double, totalOnTime As Double, totalDebtAtRisk As Double
    Dim rowStatus As String, rowSum As Double
    Dim percentDelayed As Double, percentOnTime As Double, percentDebtAtRisk As Double

    Set objWorkbook = ThisWorkbook
    Set objSheet = objWorkbook.Sheets(1)

    On Error Resume Next
    Set objSummarySheet = objWorkbook.Sheets("Summary")
    On Error GoTo 0
    
    If objSummarySheet Is Nothing Then
        Set objSummarySheet = objWorkbook.Sheets.Add(After:=objWorkbook.Sheets(objWorkbook.Sheets.Count))
        objSummarySheet.Name = "Summary"
    Else
        objSummarySheet.Cells.Clear
    End If

    lastRow = objSheet.Cells(objSheet.Rows.Count, 1).End(xlUp).Row
    grandTotal = objSheet.Cells(lastRow, 11).Value

    For i = 2 To lastRow
        rowStatus = objSheet.Cells(i, 10).Value
        rowSum = objSheet.Cells(i, 11).Value
        
        Select Case rowStatus
            Case "Delayed"
                totalDelayed = totalDelayed + rowSum
            Case "On time"
                totalOnTime = totalOnTime + rowSum
            Case "Debt at risk"
                totalDebtAtRisk = totalDebtAtRisk + rowSum
        End Select
    Next i

    percentDelayed = (totalDelayed / grandTotal) * 100
    percentOnTime = (totalOnTime / grandTotal) * 100
    percentDebtAtRisk = (totalDebtAtRisk / grandTotal) * 100

    With objSummarySheet
        .Cells(1, 1).Value = "Relative Percentage"
        .Cells(1, 2).Value = "Sum"
        .Cells(1, 3).Value = "Status"

        .Cells(2, 1).Value = Format(percentDelayed, "0.00%")
        .Cells(2, 2).Value = totalDelayed
        .Cells(2, 3).Value = "Delayed"

        .Cells(3, 1).Value = Format(percentOnTime, "0.00%")
        .Cells(3, 2).Value = totalOnTime
        .Cells(3, 3).Value = "On time"

        .Cells(4, 1).Value = Format(percentDebtAtRisk, "0.00%")
        .Cells(4, 2).Value = totalDebtAtRisk
        .Cells(4, 3).Value = "Debt at risk"

        .Cells(5, 1).Value = "100%"
        .Cells(5, 2).Value = grandTotal
        .Cells(5, 3).Value = "grandTotal"
    End With

    ' Apply formatting for the summary table
    call ApplySummaryTableStyling(objSummarySheet)
    call AutoFitColumns(objSummarySheet)
    Call CreatePieChart(objSummarySheet) 
End Sub

Sub ApplySummaryTableStyling(objSummarySheet As Object)
    ' Apply styling for the summary table in the "Summary" sheet
    Dim lastRow As Long
    lastRow = objSummarySheet.Cells(objSummarySheet.Rows.Count, 1).End(-4162).Row ' Get last row

    ' Apply font and background color for the first three cells based on the "Status" column
    Dim i As Long
    For i = 2 To lastRow
        ' Set font for all cells in the row to "David"
        objSummarySheet.Rows(i).Font.Name = "David"

        ' Apply background color based on the "Status" column (column 3) only for the first 3 columns
        If objSummarySheet.Cells(i, 3).Value <> "" Then ' Only if there is a value in the "Status" column
            Select Case objSummarySheet.Cells(i, 3).Value ' Column 3 is "Status"
                Case "Delayed"
                    objSummarySheet.Cells(i, 1).Interior.Color = RGB(255, 255, 0) ' Yellow for Delayed
                    objSummarySheet.Cells(i, 2).Interior.Color = RGB(255, 255, 0) ' Yellow for Delayed
                    objSummarySheet.Cells(i, 3).Interior.Color = RGB(255, 255, 0) ' Yellow for Delayed
                Case "On time"
                    objSummarySheet.Cells(i, 1).Interior.Color = RGB(0, 255, 0) ' Green for On time
                    objSummarySheet.Cells(i, 2).Interior.Color = RGB(0, 255, 0) ' Green for On time
                    objSummarySheet.Cells(i, 3).Interior.Color = RGB(0, 255, 0) ' Green for On time
                Case "Debt at risk"
                    objSummarySheet.Cells(i, 1).Interior.Color = RGB(255, 182, 193) ' Pink for Debt at risk
                    objSummarySheet.Cells(i, 2).Interior.Color = RGB(255, 182, 193) ' Pink for Debt at risk
                    objSummarySheet.Cells(i, 3).Interior.Color = RGB(255, 182, 193) ' Pink for Debt at risk
            End Select
        End If
    Next i

    ' Apply background color to the Grand Total row (last row)
    objSummarySheet.Cells(lastRow, 1).Interior.Color = RGB(192, 192, 192) ' Gray for Grand Total row
    objSummarySheet.Cells(lastRow, 2).Interior.Color = RGB(192, 192, 192) ' Gray for Grand Total row
    objSummarySheet.Cells(lastRow, 3).Interior.Color = RGB(192, 192, 192) ' Gray for Grand Total row

    ' Apply borders to the summary table (first 3 columns)
    With objSummarySheet.Range("A1:C" & lastRow) ' Use only up to the last actual data row
        .Borders.LineStyle = 1 ' xlContinuous
        .Borders.Color = RGB(0, 0, 0) ' Black border color
    End With

    ' Center-align text in the summary table
    With objSummarySheet.Range("A1:C" & lastRow)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Format the Sum column (Column B) to include commas
    objSummarySheet.Columns("B:B").NumberFormat = "#,##0"
End Sub

Sub AutoFitColumns(objSummarySheet As Object)
    ' AutoFit all columns in the Summary sheet
    objSummarySheet.Columns("A:C").AutoFit
End Sub

Sub CreatePieChart(objSummarySheet As Object)
    Dim objChart As Object
    Dim dataRange As Range
    Dim chartSheet As Worksheet

    ' Create a new sheet called "Chart"
    On Error Resume Next
    Set chartSheet = ThisWorkbook.Sheets("Chart")
    On Error GoTo 0

    ' If the "Chart" sheet already exists, clear it; otherwise, create a new one
    If chartSheet Is Nothing Then
        Set chartSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        chartSheet.Name = "Chart"
    Else
        chartSheet.Cells.Clear ' Clear the existing chart if the sheet already exists
    End If

    ' Define the data range for the pie chart (the second and third columns in the summary table)
    Set dataRange = objSummarySheet.Range("B2:B4") ' Sum column (from row 2 to row 4)

    ' Add a chart to the "Chart" sheet (positioned with a specific size)
    Set objChart = chartSheet.ChartObjects.Add(Left:=100, Top:=50, Width:=375, Height:=225)

    ' Set the chart type to pie chart
    objChart.Chart.ChartType = xlPie

    ' Set the chart data source (categories and their corresponding values)
    objChart.Chart.SetSourceData Source:=dataRange

    ' Set the chart's category labels (the "Status" column)
    objChart.Chart.SeriesCollection(1).XValues = objSummarySheet.Range("C2:C4") ' Status column

    ' Set the pie chart colors (Pink, Green, Yellow)
    objChart.Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(255, 182, 193) ' Pink
    objChart.Chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = RGB(0, 255, 0) ' Green
    objChart.Chart.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = RGB(255, 255, 0) ' Yellow

    ' Optional: Apply data labels to the pie chart
    objChart.Chart.ApplyDataLabels
End Sub
