Option Explicit

'Deletes sheets that are not in the desired array of sheets
Sub DeleteTempSheets(desiredSheets() As String)
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim i As Long, sheetName As String
    Dim objSheet As Worksheet
    Dim sheetExists As Boolean
    Dim sheet As Variant ' Declare the sheet variable as a String for the For Each loop
    
    ' Check if workbook has at least one sheet
    If ThisWorkbook.Sheets.Count < 1 Then
        MsgBox "Error: Workbook must contain at least one sheet.", vbCritical
        Exit Sub
    End If
    
    ' Loop through each sheet in the workbook
    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Set objSheet = ThisWorkbook.Sheets(i)
        sheetName = objSheet.Name
        
        ' Check if the sheet is in the desiredSheets array
        sheetExists = False
        For Each sheet In desiredSheets  ' Loop through the array of strings
            If sheetName = sheet Then
                sheetExists = True
                Exit For
            End If
        Next sheet
        
        ' If sheet is not in the desiredSheets array, delete it
        If Not sheetExists Then
            objSheet.Delete
        End If
    Next i
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
