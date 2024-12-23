Option Explicit

Sub SplitDataByManagerWithTotals()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long
    Dim managerName As String, currentManager As String
    Dim managerSheet As Worksheet
    Dim dict As Object
    Dim headerRow As Range
    Dim dataRange As Range
    Dim managerTotal As Double
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
    
    ' Create dictionary to track manager data
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' First pass: Collect all data per manager
    currentManager = ""
    managerTotal = 0
    
    For i = 2 To lastRow
        managerName = Trim(ws.Cells(i, 1).Value)
        
        ' Skip empty rows
        If managerName <> "" Then
            ' If we haven't seen this manager before, create new dictionary entry
            If Not dict.exists(managerName) Then
                Set dict(managerName) = CreateObject("Scripting.Dictionary")
                dict(managerName)("Rows") = CreateObject("Scripting.Dictionary")
                dict(managerName)("Total") = 0
            End If
            
            ' Store row data
            Dim rowData As Range
            Set rowData = ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol))
            dict(managerName)("Rows").Add dict(managerName)("Rows").Count + 1, rowData.Value
            
            ' Add to manager total if last column contains a number
            If IsNumeric(ws.Cells(i, lastCol).Value) Then
                dict(managerName)("Total") = dict(managerName)("Total") + ws.Cells(i, lastCol).Value
            End If
        End If
    Next i
    
    ' Second pass: Create sheets for each manager
    Dim grandTotal As Double
    grandTotal = 0
    
    For Each managerName In dict.keys
        ' Delete sheet if it already exists
        Application.DisplayAlerts = False
        On Error Resume Next
        ThisWorkbook.Sheets(managerName).Delete
        On Error GoTo ErrorHandler
        Application.DisplayAlerts = True
        
        ' Create new sheet
        Set managerSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        managerSheet.Name = managerName
        
        ' Copy header row
        headerRow.Copy managerSheet.Range("A1")
        
        ' Write data rows
        nextRow = 2
        Dim rowNum As Variant
        For Each rowNum In dict(managerName)("Rows").keys
            Dim rowValues As Variant
            rowValues = dict(managerName)("Rows")(rowNum)
            
            ' Write row values
            With managerSheet
                Dim col As Long
                For col = 1 To UBound(rowValues, 2)
                    .Cells(nextRow, col).Value = rowValues(1, col)
                Next col
                
                ' Format date columns (7, 8, 9)
                If IsDate(.Cells(nextRow, 7).Value) Then .Cells(nextRow, 7).NumberFormat = "dd/mm/yyyy"
                If IsDate(.Cells(nextRow, 8).Value) Then .Cells(nextRow, 8).NumberFormat = "dd/mm/yyyy"
                If IsDate(.Cells(nextRow, 9).Value) Then .Cells(nextRow, 9).NumberFormat = "dd/mm/yyyy"
                
                ' Format amount column
                .Cells(nextRow, lastCol).NumberFormat = "#,##0"
            End With
            
            nextRow = nextRow + 1
        Next rowNum
        
        ' Add total row
        With managerSheet
            .Cells(nextRow, 3).Value = managerName & " Total"
            .Cells(nextRow, lastCol).Value = dict(managerName)("Total")
            .Cells(nextRow, lastCol).NumberFormat = "#,##0"
            .Range(.Cells(nextRow, 1), .Cells(nextRow, lastCol)).Interior.Color = RGB(240, 240, 240)
            .Range(.Cells(nextRow, 1), .Cells(nextRow, lastCol)).Font.Bold = True
        End With
        
        grandTotal = grandTotal + dict(managerName)("Total")
        nextRow = nextRow + 1
        
        ' Format sheet
        With managerSheet
            ' Add grand total
            .Cells(nextRow, 1).Value = "Grand Total"
            .Cells(nextRow, lastCol).Value = grandTotal
            .Cells(nextRow, lastCol).NumberFormat = "#,##0"
            .Range(.Cells(nextRow, 1), .Cells(nextRow, lastCol)).Interior.Color = RGB(200, 200, 200)
            .Range(.Cells(nextRow, 1), .Cells(nextRow, lastCol)).Font.Bold = True
            
            ' Format header
            .Range("A1:K1").Font.Bold = True
            .Range("A1:K1").Interior.Color = RGB(220, 220, 220)
            
            ' General formatting
            .UsedRange.Columns.AutoFit
            .UsedRange.Borders.LineStyle = xlContinuous
            .DisplayRightToLeft = True
        End With
    Next managerName
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = True
    MsgBox "Process complete! Created " & dict.Count & " sheets.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub