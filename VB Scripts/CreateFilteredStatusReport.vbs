Sub CreateFilteredStatusReport()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Initialize variables
    Dim sourceSheet As Worksheet
    Dim filteredSheet As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim statusCol As Long
    Dim rowToDelete As Long
    Dim currentTotal As Double
    
    ' Reference the source sheet (Collection Report)
    Set sourceSheet = ThisWorkbook.Sheets("Collection Report")
    
    ' Create new sheet for filtered data
    Dim sheetName As String
    sheetName = "Delayed Payments Report"
    
    ' Add new sheet and copy all data
    Set filteredSheet = ThisWorkbook.Sheets.Add
    filteredSheet.Name = sheetName
    
    sourceSheet.UsedRange.Copy
    filteredSheet.Range("A1").PasteSpecial xlPasteValues
    filteredSheet.Range("A1").PasteSpecial xlPasteFormats
    
    ' Get last row and column
    lastRow = filteredSheet.Cells(filteredSheet.Rows.Count, 1).End(xlUp).Row
    lastCol = filteredSheet.Cells(1, filteredSheet.Columns.Count).End(xlToLeft).Column
    
    ' Find the Status column
    statusCol = 0
    For i = 1 To lastCol
        If filteredSheet.Cells(1, i).Value = "Status" Then
            statusCol = i
            Exit For
        End If
    Next i
    
    ' Remove rows with "On time" status and recalculate totals
    Dim balanceCol As Long
    balanceCol = lastCol ' Balance is the last column
    
    rowToDelete = 2 ' Start from first data row
    Do While rowToDelete <= lastRow
        If filteredSheet.Cells(rowToDelete, 3).Value = "Total" Then
            ' Recalculate total
            currentTotal = 0
            Dim startRow As Long
            startRow = rowToDelete - 1
            
            ' Go backwards until we find the start of this group
            Do While startRow >= 2 And filteredSheet.Cells(startRow, 3).Value <> "Total"
                If filteredSheet.Cells(startRow, statusCol).Value <> "" Then ' Skip already deleted rows
                    currentTotal = currentTotal + filteredSheet.Cells(startRow, balanceCol).Value
                End If
                startRow = startRow - 1
            Loop
            
            ' Update total
            filteredSheet.Cells(rowToDelete, balanceCol).Value = currentTotal
            rowToDelete = rowToDelete + 1
            
        ElseIf filteredSheet.Cells(rowToDelete, 1).Value = "Grand Total" Then
            ' Recalculate grand total
            currentTotal = 0
            For i = 2 To rowToDelete - 1
                If filteredSheet.Cells(i, 3).Value = "Total" Then
                    currentTotal = currentTotal + filteredSheet.Cells(i, balanceCol).Value
                End If
            Next i
            filteredSheet.Cells(rowToDelete, balanceCol).Value = currentTotal
            rowToDelete = rowToDelete + 1
            
        ElseIf filteredSheet.Cells(rowToDelete, statusCol).Value = "On time" Then
            ' Delete "On time" rows
            filteredSheet.Rows(rowToDelete).Delete
            lastRow = lastRow - 1
            
        Else
            rowToDelete = rowToDelete + 1
        End If
    Loop
    
    ' Reformat the filtered sheet
    With filteredSheet
        .UsedRange.Columns.AutoFit
        .UsedRange.Borders.LineStyle = xlContinuous
    End With
    
CleanExit:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub