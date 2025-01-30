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

    ' Find the chart sheet
    On Error Resume Next
    Set wsChart = ThisWorkbook.Sheets("Chart")
    If wsChart Is Nothing Then
        MsgBox "Error: 'Chart' sheet not found.", vbCritical
        Exit Sub
    End If
    Set objChart = wsChart.ChartObjects(1) ' Adjust index if multiple charts exist
    On Error GoTo 0

    ' Ensure chart exists
    If objChart Is Nothing Then
        MsgBox "Error: No chart found in 'Chart' sheet.", vbCritical
        Exit Sub
    End If

    ' ✅ Remove background and make it transparent
    With objChart.Chart
        ' Remove ChartArea background (fully transparent)
        .ChartArea.Format.Fill.Visible = msoFalse
        .ChartArea.Format.Line.Visible = msoFalse

        ' Remove PlotArea background (fully transparent)
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
    End With

    ' Export the sheet to HTML
    ws.SaveAs Filename:=htmlFile, FileFormat:=xlHTML

    ' Make the chart visible again (if it was hidden)
    If Not objChart Is Nothing Then
        objChart.Visible = True
    End If
End Sub
