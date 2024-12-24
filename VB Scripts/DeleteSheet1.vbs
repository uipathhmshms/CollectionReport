Option Explicit

'Deletes the sheet named "Sheet1" if it exists
Sub DeleteSheet1()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim i, sheetName, objSheet
    
    ' Check if workbook has at least one sheet
    If ThisWorkbook.Sheets.Count < 1 Then
        MsgBox "Error: Workbook must contain at least one sheet.", vbCritical
        Exit Sub
    End If
    
    ' Loop through each sheet in the workbook
    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Set objSheet = ThisWorkbook.Sheets(i)
        sheetName = objSheet.Name

        ' Check if the sheet is named "Sheet1" and delete it
        If sheetName = "Sheet1" Then
            objSheet.Delete
            Exit For ' Exit the loop once the sheet is deleted
        End If
    Next
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
