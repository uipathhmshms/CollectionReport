Option Explicit

'Deletes all opned sheets except the last one
Sub DeletePreviousSheets()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
	Dim i, lastSheetIndex, sheetName, objSheet
	
    ' Check if workbook has at least one sheet
    If ThisWorkbook.Sheets.Count < 1 Then
        MsgBox "Error: Workbook must contain at least one sheet.", vbCritical
        Exit Sub
    End If
    
	' Get the index of the last sheet in the workbook
	lastSheetIndex = ThisWorkbook.Sheets.Count

	' Loop through each sheet in the workbook (start from the second-to-last sheet to avoid skipping)
	For i = lastSheetIndex - 1 To 1 Step -1
		Set objSheet = ThisWorkbook.Sheets(i)
		sheetName = objSheet.Name

		' Delete the sheet
		objSheet.Delete
	Next
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
