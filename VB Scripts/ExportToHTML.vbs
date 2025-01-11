Sub ExportToHTML()
	Application.ScreenUpdating = False
	
    Dim ws As Worksheet
    Dim htmlFile As String
    Dim fileNameWithoutExt As String
    Dim objChart As Object

    ' Set the sheet to be exported
    Set ws = ThisWorkbook.Sheets("Summary") 
	
	If ws Is Nothing Then
        MsgBox "Summary sheet not found. Please create a summary first.", vbCritical
        Exit Sub
    End If
	
    ' Get the name of the workbook without the extension
    fileNameWithoutExt = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    ' Define the output HTML file path using the workbook name
    htmlFile = ThisWorkbook.Path & "\" & fileNameWithoutExt & ".html"  

    ' Export the sheet to HTML
    ws.SaveAs Filename:=htmlFile, FileFormat:=xlHTML

End Sub
