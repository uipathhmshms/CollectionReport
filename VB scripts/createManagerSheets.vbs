Option Explicit

Sub SplitDataByManager()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim managerName As String
    Dim managerSheet As Worksheet
    Dim dict As Object
    Dim dataRange As Range
    Dim headerRow As Range
    Dim rowToCopy As Range
    Dim nextRow As Long
    
    ' Set reference to first worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Find last row and column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Exit if no data
    If lastRow <= 1 Then
        MsgBox "No data found in the first sheet!", vbExclamation
        Exit Sub
    End If
    
    ' Store header row
    Set headerRow = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    
    ' Create dictionary to track manager sheets
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop through each row
    For i = 2 To lastRow
        ' Get manager name from first column
        managerName = Trim(ws.Cells(i, 1).Value)
        
        ' Skip if manager name is empty
        If managerName = "" Then Continue For
        
        ' Create new sheet for manager if it doesn't exist
        If Not dict.exists(managerName) Then
            ' Delete sheet if it already exists
            Application.DisplayAlerts = False
            On Error Resume Next
            ThisWorkbook.Sheets(managerName).Delete
            On Error GoTo ErrorHandler
            Application.DisplayAlerts = True
            
            ' Create new sheet
            Set managerSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            managerSheet.Name = managerName
            
            ' Copy header row to new sheet
            headerRow.Copy managerSheet.Range("A1")
            
            ' Format header row
            With managerSheet.Range("A1").Resize(1, lastCol)
                .Font.Bold = True
                .Interior.Color = RGB(220, 220, 220)
                .Borders.LineStyle = xlContinuous
            End With
            
            ' Add sheet to dictionary
            dict.Add managerName, 2 ' 2 is the next row to write to
        End If
        
        ' Get reference to manager's sheet and next row to write to
        Set managerSheet = ThisWorkbook.Sheets(managerName)
        nextRow = dict(managerName)
        
        ' Copy entire row to manager's sheet
        Set rowToCopy = ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol))
        rowToCopy.Copy managerSheet.Cells(nextRow, 1)
        
        ' Update next row in dictionary
        dict(managerName) = nextRow + 1
    Next i
    
    ' Format all created sheets
    For Each managerName In dict.keys
        Set managerSheet = ThisWorkbook.Sheets(managerName)
        
        With managerSheet
            ' Autofit columns
            .UsedRange.Columns.AutoFit
            
            ' Add borders
            .UsedRange.Borders.LineStyle = xlContinuous
            
            ' Format numbers in the last column (Sum of יתרה)
            .Range(.Cells(2, lastCol), .Cells(dict(managerName) - 1, lastCol)).NumberFormat = "#,##0"
            
            ' Set RTL for Hebrew support
            .DisplayRightToLeft = True
        End With
    Next managerName
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    MsgBox "Process complete! Created " & dict.Count & " sheets.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub