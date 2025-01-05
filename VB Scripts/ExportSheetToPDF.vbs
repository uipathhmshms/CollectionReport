Sub ExportSheetToPDF()
    Dim ws As Worksheet
    Dim fileNameWithoutExt As String
    Dim pdfFile As String

    ' Set the sheet to be exported
    Set ws = ThisWorkbook.Sheets("Summary") 
	
    ' Get the name of the workbook without the extension
    fileNameWithoutExt = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    ' Define the output PDF file path using the workbook name
    pdfFile = ThisWorkbook.Path & "\" & fileNameWithoutExt & ".pdf"

    ' Export the sheet to PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFile

End Sub
