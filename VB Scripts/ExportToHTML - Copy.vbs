Sub ExportToHTML()
    Dim ws As Worksheet
    Dim htmlFile As String
    Dim fileNameWithoutExt As String

    ' Set the active worksheet
    Set ws = ThisWorkbook.Sheets(1)

    ' Get the name of the workbook without the extension
    fileNameWithoutExt = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    ' Define the output HTML file path using the workbook name
    htmlFile = ThisWorkbook.Path & "\" & fileNameWithoutExt & ".html"

    ' Export the sheet to HTML
    ws.SaveAs Filename:=htmlFile, FileFormat:=xlHTML

End Sub
