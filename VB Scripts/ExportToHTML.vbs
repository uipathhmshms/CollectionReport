Sub ExportToHTML()
    Dim ws As Worksheet
	Dim wsChart As Worksheet
    Dim htmlFile As String
    Dim fileNameWithoutExt As String
    Dim objChart As Object
	Dim lastRow as Long

    ' Set the active worksheet
    Set ws = ThisWorkbook.Sheets(1)

    ' Get the name of the workbook without the extension
    fileNameWithoutExt = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    ' Define the output HTML file path using the workbook name
    htmlFile = ThisWorkbook.Path & "\" & fileNameWithoutExt & ".html"

    ' Find and hide the chart (assuming the chart is the first chart object)
	Set wsChart= ThisWorkbook.Sheets("Chart")
    On Error Resume Next
    Set objChart = wsChart.ChartObjects(1) ' Adjust the index if there are multiple charts
    If Not objChart Is Nothing Then
        ' Remove borders from the Chart Area
        objChart.Chart.ChartArea.Format.Line.Visible = msoFalse

        ' Remove borders from the Plot Area
        objChart.Chart.PlotArea.Format.Line.Visible = msoFalse
        
        ' Set the PlotArea and ChartArea background color to black (RGB(0, 0, 0))
        objChart.Chart.PlotArea.Format.Fill.ForeColor.RGB = RGB(18, 18, 18) ' Black  
        objChart.Chart.ChartArea.Format.Fill.ForeColor.RGB = RGB(18, 18, 18) ' Black
    End If
    On Error GoTo 0

	' ' Convert the table to a chart so it will export as image and not html
	' ' Find the last row of the table (assuming table is in columns A and B)
    ' lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
	' ' Create a chart from the table range (for example, data in columns A and B)
    ' Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=50, Width:=400, Height:=300)
    ' ' Set the chart's data range (e.g., columns A and B)
    ' chartObj.Chart.SetSourceData Source:=ws.Range("A1:B" & lastRow)
    ' ' Set the chart type (for example, column chart)
    ' chartObj.Chart.ChartType = xlColumnClustered

    ' Export the sheet to HTML
    ws.SaveAs Filename:=htmlFile, FileFormat:=xlHTML

    ' Make the chart visible again (if it was hidden)
    If Not objChart Is Nothing Then
        objChart.Visible = True
    End If
End Sub
