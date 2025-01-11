Sub CreatePieChart()
    Dim objChart As Object
    Dim dataRange As Range
	Dim objWorkbook As Workbook
    Dim chartSheet As Worksheet
	Dim objSummarySheet As Worksheet
	
    Set objWorkbook = ThisWorkbook  
    Set objSummarySheet = objWorkbook.Sheets("Summary")  
	
    ' Create a new sheet called "Chart"
    On Error Resume Next
    Set chartSheet = ThisWorkbook.Sheets("Chart")
    On Error GoTo 0

    ' If the "Chart" sheet already exists, clear it; otherwise, create a new one
    If chartSheet Is Nothing Then
        Set chartSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        chartSheet.Name = "Chart"
    Else
        chartSheet.Cells.Clear
    End If

    ' Define the data range for the pie chart (the second and third columns in the summary table)
    Set dataRange = objSummarySheet.Range("B2:B4") ' Sum column (from row 2 to row 4)

    ' Add a chart to the "Chart" sheet
    Set objChart = chartSheet.ChartObjects.Add(Left:=100, Top:=50, Width:=375, Height:=225)

    ' Set the chart type to pie chart
    objChart.Chart.ChartType = xlPie

    ' Set the chart data source
    objChart.Chart.SetSourceData Source:=dataRange

    ' Remove the legend
    objChart.Chart.HasLegend = False

    ' Set the pie chart colors to match the table (Green, Yellow, Pink)
    objChart.Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(0, 255, 0) ' Green
    objChart.Chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = RGB(255, 255, 0) ' Yellow
    objChart.Chart.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = RGB(255, 182, 193) ' Pink

    ' Remove data labels
    On Error Resume Next
    objChart.Chart.SeriesCollection(1).DataLabels.Delete
    On Error GoTo 0
End Sub