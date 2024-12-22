' Create a new Class Module named "RowData" and put this code in it:
' --- Start of RowData Class Module ---

Public AccountName As String
Public BudgetSectionExtra As String
Public Reference As String
Public ReferenceDate As Date
Public PaymentDueDate As Date
Public ActivityDate As Date
Public Balance As Double
' --- End of RowData Class Module ---


Sub TransformData()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim inputSheet As Worksheet
    Dim outputSheet As Worksheet
    Dim inputLastRow As Long
    Dim outputRow As Long
    Dim groupedData As Object
    Dim grandTotal As Double
    
    ' Initialize worksheets
    Set inputSheet = ThisWorkbook.Sheets(1)
    
    ' Check if the input sheet has data
    inputLastRow = GetLastRow(inputSheet)
    If inputLastRow < 2 Then
        MsgBox "Error: No data found in input sheet.", vbCritical
        GoTo CleanExit
    End If
    
    Set outputSheet = CreateOutputSheet()
    If outputSheet Is Nothing Then
        MsgBox "Error: Could not create output sheet.", vbCritical
        GoTo CleanExit
    End If
    
    ' Copy the header and initialize processing
    CopyHeader inputSheet, outputSheet
    Set groupedData = CreateObject("Scripting.Dictionary")
    outputRow = 2
    grandTotal = 0
    
    ' Process input data and group it
    ProcessInputData inputSheet, groupedData, inputLastRow
    
    ' Write grouped data and totals to output sheet
    WriteDataToOutputSheet groupedData, outputSheet, outputRow, grandTotal
    
    ' Final formatting and cleanup
    ApplyFinalFormatting outputSheet
    
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    If Err.Number = 0 Then
        MsgBox "Transformation complete. Data has been written to sheet '" & outputSheet.Name & "'.", vbInformation
    End If
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    Resume CleanExit
End Sub

Private Function CreateOutputSheet() As Worksheet
    Dim sheetName As String
    Dim i As Integer
    
    i = 1
    sheetName = "Transformed_Data"
    
    ' Handle existing sheets with similar names
    Do While SheetExists(sheetName & IIf(i = 1, "", "_" & i))
        i = i + 1
        If i > 100 Then Exit Function  ' Prevent infinite loop
    Loop
    
    sheetName = sheetName & IIf(i = 1, "", "_" & i)
    Set CreateOutputSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    CreateOutputSheet.Name = sheetName
End Function

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function GetLastRow(sheet As Worksheet) As Long
    With sheet
        GetLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
End Function

Private Sub CopyHeader(inputSheet As Worksheet, outputSheet As Worksheet)
    With outputSheet
        inputSheet.Rows(1).Copy
        .Rows(1).PasteSpecial xlPasteValues
        .Rows(1).PasteSpecial xlPasteFormats
        .Cells(1, .Cells(1, .Columns.Count).End(xlToLeft).Column + 1).Value = "Status"
        Application.CutCopyMode = False
    End With
End Sub

Private Sub ProcessInputData(inputSheet As Worksheet, ByRef groupedData As Object, inputLastRow As Long)
    Dim dataArr As Variant
    Dim i As Long
    Dim groupKey As String
    Dim subGroupKey As String
    Dim rowData As RowData
    
    ' Read data into an array for faster processing
    dataArr = inputSheet.Range("A2:J" & inputLastRow).Value2
    
    For i = 1 To UBound(dataArr)
        ' Create group keys
        groupKey = CStr(dataArr(i, 1)) & "|" & CStr(dataArr(i, 2))  ' Project Manager | Account Code
        subGroupKey = CStr(dataArr(i, 3)) & "|" & CStr(dataArr(i, 4))  ' Account Name | Budget Section
        
        ' Initialize group dictionaries if they don't exist
        If Not groupedData.Exists(groupKey) Then
            groupedData.Add groupKey, CreateObject("Scripting.Dictionary")
        End If
        
        If Not groupedData(groupKey).Exists(subGroupKey) Then
            groupedData(groupKey).Add subGroupKey, CreateObject("Scripting.Dictionary")
        End If
        
        ' Populate row data
        Set rowData = New RowData
        With rowData
            .AccountName = CStr(dataArr(i, 3))
            .BudgetSectionExtra = CStr(dataArr(i, 5))
            .Reference = CStr(dataArr(i, 6))
            On Error Resume Next
            .ReferenceDate = CDate(dataArr(i, 7))
            .PaymentDueDate = CDate(dataArr(i, 8))
            .ActivityDate = CDate(dataArr(i, 9))
            On Error GoTo 0
            .Balance = IIf(IsNumeric(dataArr(i, 10)), CDbl(dataArr(i, 10)), 0)
        End With
        
        ' Add to grouped data
        groupedData(groupKey)(subGroupKey).Add _
            groupedData(groupKey)(subGroupKey).Count, rowData
    Next i
End Sub

Private Sub WriteDataToOutputSheet(ByRef groupedData As Object, outputSheet As Worksheet, _
                                 ByRef outputRow As Long, ByRef grandTotal As Double)
    Dim groupKey As Variant, subGroupKey As Variant, rowKey As Variant
    Dim splitGroup() As String, splitSubGroup() As String
    Dim lastGroupKey As String
    Dim groupTotal As Double
    Dim rowData As RowData
    
    lastGroupKey = ""
    
    For Each groupKey In groupedData.Keys
        groupTotal = 0
        splitGroup = Split(groupKey, "|")
        
        For Each subGroupKey In groupedData(groupKey).Keys
            splitSubGroup = Split(subGroupKey, "|")
            
            For Each rowKey In groupedData(groupKey)(subGroupKey).Keys
                Set rowData = groupedData(groupKey)(subGroupKey)(rowKey)
                
                With outputSheet
                    ' Write group identifiers only if new group
                    If groupKey <> lastGroupKey Then
                        .Cells(outputRow, 1).Value = splitGroup(0)  ' Project Manager
                        .Cells(outputRow, 2).Value = splitGroup(1)  ' Account Code
                        .Cells(outputRow, 3).Value = splitSubGroup(0)  ' Account Name
                    End If
                    
                    ' Write row data
                    WriteRowData outputSheet, outputRow, rowData, splitSubGroup(1)  ' Budget Section
                    
                    groupTotal = groupTotal + rowData.Balance
                    outputRow = outputRow + 1
                End With
                lastGroupKey = groupKey
            Next rowKey
        Next subGroupKey
        
        ' Add group total
        WriteGroupTotal outputSheet, outputRow, groupTotal, splitGroup(0)
        outputRow = outputRow + 1
        grandTotal = grandTotal + groupTotal
        lastGroupKey = ""
    Next groupKey
    
    ' Add grand total
    WriteGrandTotal outputSheet, outputRow, grandTotal
End Sub

Private Sub WriteRowData(ws As Worksheet, row As Long, rowData As RowData, budgetSection As String)
    With ws
        .Cells(row, 4).Value = budgetSection
        .Cells(row, 5).Value = rowData.BudgetSectionExtra
        .Cells(row, 6).Value = rowData.Reference
        .Cells(row, 7).Value = rowData.ReferenceDate
        .Cells(row, 8).Value = rowData.PaymentDueDate
        .Cells(row, 9).Value = rowData.ActivityDate
        .Cells(row, 10).Value = GetStatus(rowData.PaymentDueDate)
        .Cells(row, 11).Value = rowData.Balance
        .Cells(row, 11).NumberFormat = "#,##0"
    End With
End Sub

Private Function GetStatus(paymentDate As Date) As String
    If paymentDate = 0 Then
        GetStatus = "Invalid date"
    ElseIf paymentDate < Date Then
        GetStatus = "Delayed"
    ElseIf paymentDate > Date Then
        GetStatus = "On time"
    Else
        GetStatus = "Due today"
    End If
End Function

Private Sub WriteGroupTotal(ws As Worksheet, row As Long, total As Double, projectManager As String)
    With ws
        .Cells(row, 3).Value = projectManager & " Total"
        .Cells(row, 11).Value = total
        .Range(.Cells(row, 1), .Cells(row, 11)).Interior.Color = RGB(240, 240, 240)
        .Cells(row, 11).Font.Bold = True
        .Cells(row, 11).NumberFormat = "#,##0"
    End With
End Sub

Private Sub WriteGrandTotal(ws As Worksheet, row As Long, total As Double)
    With ws
        .Cells(row, 1).Value = "Grand Total"
        .Cells(row, 11).Value = total
        .Range(.Cells(row, 1), .Cells(row, 11)).Interior.Color = RGB(200, 200, 200)
        .Range(.Cells(row, 1), .Cells(row, 11)).Font.Bold = True
        .Cells(row, 11).NumberFormat = "#,##0"
    End With
End Sub

Private Sub ApplyFinalFormatting(ws As Worksheet)
    With ws.UsedRange
        .Columns.AutoFit
        .Borders.LineStyle = xlContinuous
        .Rows(1).Font.Bold = True
    End With
End Sub