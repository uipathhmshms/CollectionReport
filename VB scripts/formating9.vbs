Option Explicit

Sub TransformData()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim inputSheet As Worksheet
    Dim outputSheet As Worksheet
    Dim summarySheet As Worksheet ' New summary sheet for grand totals
    Dim inputLastRow As Long
    Dim outputRow As Long
    Dim groupedData As Object
    Dim grandTotal As Double
    Dim today As Date
    Dim status As String
    Dim groupKey As Variant
    Dim subGroupKey As Variant
    Dim rowKey As Variant
    Dim rowData As Object
    Dim groupTotal As Double
    Dim lastGroupKey As String
    Dim projectManager As String, accountCode As String, accountName As String
    Dim budgetSection As String, budgetSectionExtra As String
    Dim reference As String, referenceDate As Date
    Dim paymentDate As Date, activityDate As Date
    Dim balance As Double

    ' Check if workbook has at least one sheet
    If ThisWorkbook.Sheets.Count < 1 Then
        MsgBox "Error: Workbook must contain at least one sheet.", vbCritical
        Exit Sub
    End If
    
    ' Reference the active workbook and sheets
    Set inputSheet = ThisWorkbook.Sheets(1) ' Input data is on the first sheet
    
    ' Check if input sheet has data
    If Application.WorksheetFunction.CountA(inputSheet.Cells) = 0 Then
        MsgBox "Error: Input sheet is empty.", vbCritical
        Exit Sub
    End If
    
    ' Create new sheet for each manager
    Dim sheetName As String
    sheetName = "Transformed_Data_"
    
    ' Create a summary sheet to store grand totals of each manager
    Set summarySheet = ThisWorkbook.Sheets.Add
    summarySheet.Name = "Manager_Summary"
    summarySheet.Cells(1, 1).Value = "Manager"
    summarySheet.Cells(1, 2).Value = "Grand Total"
    
    ' Initialize the row for summary sheet
    Dim summaryRow As Long
    summaryRow = 2
    
    ' Get the last row in the input sheet
    inputLastRow = inputSheet.Cells(inputSheet.Rows.Count, 1).End(xlUp).Row
    
    If inputLastRow < 2 Then
        MsgBox "Error: No data found in input sheet.", vbCritical
        Exit Sub
    End If
    
    ' Initialize variables
    Set groupedData = CreateObject("Scripting.Dictionary")
    grandTotal = 0
    outputRow = 2 ' Start writing from the second row
    
    ' Group data by "שם מנהל" and then "קוד חשבון" and "שם חשבון"
    Dim inputRow As Long
    Dim dataRange As Range
    Set dataRange = inputSheet.Range("A2:J" & inputLastRow)
    
    ' Read data into array for faster processing
    Dim dataArr As Variant
    dataArr = dataRange.Value
    
    For inputRow = 1 To UBound(dataArr, 1)
        projectManager = CStr(dataArr(inputRow, 1))
        accountCode = CStr(dataArr(inputRow, 2))
        accountName = CStr(dataArr(inputRow, 3))
        budgetSection = CStr(dataArr(inputRow, 4))
        budgetSectionExtra = CStr(dataArr(inputRow, 5))
        reference = CStr(dataArr(inputRow, 6))
        
        ' Handle potential date conversion errors
        On Error Resume Next
        referenceDate = CDate(dataArr(inputRow, 7))
        paymentDate = CDate(dataArr(inputRow, 8))
        activityDate = CDate(dataArr(inputRow, 9))
        On Error GoTo ErrorHandler
        
        If IsNumeric(dataArr(inputRow, 10)) Then
            balance = CDbl(dataArr(inputRow, 10))
        Else
            balance = 0
        End If
        
        ' Create group key as "שם מנהל|קוד חשבון"
        groupKey = projectManager & "|" & accountCode
        
        ' Add group to the dictionary if it doesn't exist
        If Not groupedData.Exists(groupKey) Then
            groupedData.Add groupKey, CreateObject("Scripting.Dictionary")
        End If
        
        ' Create subgroup key as "שם חשבון|קטע תקציב"
        subGroupKey = accountName & "|" & budgetSection
        If Not groupedData(groupKey).Exists(subGroupKey) Then
            groupedData(groupKey).Add subGroupKey, CreateObject("Scripting.Dictionary")
        End If
        
        ' Add row data to the subgroup dictionary
        Set rowData = CreateObject("Scripting.Dictionary")
        With rowData
            .Add "AccountName", accountName
            .Add "BudgetSectionExtra", budgetSectionExtra
            .Add "Reference", reference
            .Add "ReferenceDate", referenceDate
            .Add "PaymentDueDate", paymentDate
            .Add "ActivityDate", activityDate
            .Add "Balance", balance
        End With
        
        groupedData(groupKey)(subGroupKey).Add groupedData(groupKey)(subGroupKey).Count, rowData
    Next inputRow
    
    ' Process each group
    today = Date
    lastGroupKey = ""
    
    For Each groupKey In groupedData.Keys
        groupTotal = 0
        Dim splitGroup() As String
        splitGroup = Split(CStr(groupKey), "|")
        projectManager = splitGroup(0)
        accountCode = splitGroup(1)
        
        ' Create a new sheet for each manager if not already created
        If Not WorksheetExists(projectManager) Then
            Set outputSheet = ThisWorkbook.Sheets.Add
            outputSheet.Name = projectManager
            outputSheet.Rows(1).Value = inputSheet.Rows(1).Value
            outputSheet.Cells(1, outputSheet.Columns.Count).End(xlToLeft).Offset(0, 1).Value = "Status"
        Else
            Set outputSheet = ThisWorkbook.Sheets(projectManager)
        End If
        
        outputRow = outputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row + 1
        
        For Each subGroupKey In groupedData(groupKey).Keys
            For Each rowKey In groupedData(groupKey)(subGroupKey).Keys
                Set rowData = groupedData(groupKey)(subGroupKey)(rowKey)
                
                paymentDate = rowData("PaymentDueDate")
                
                ' Determine status
                If IsDate(paymentDate) Then
                    If paymentDate < today Then
                        status = "Delayed"
                    ElseIf paymentDate > today Then
                        status = "On time"
                    Else
                        status = "Due today"
                    End If
                Else
                    status = "Invalid date"
                End If
                
                ' Write row to the output sheet for each manager
                With outputSheet
                    .Cells(outputRow, 1).Value = projectManager
                    .Cells(outputRow, 2).Value = accountCode
                    .Cells(outputRow, 3).Value = rowData("AccountName")
                    .Cells(outputRow, 4).Value = Split(subGroupKey, "|")(1)
                    .Cells(outputRow, 5).Value = rowData("BudgetSectionExtra")
                    .Cells(outputRow, 6).Value = rowData("Reference")
                    .Cells(outputRow, 7).Value = rowData("ReferenceDate")
                    .Cells(outputRow, 8).Value = rowData("PaymentDueDate")
                    .Cells(outputRow, 9).Value = rowData("ActivityDate")
                    .Cells(outputRow, 10).Value = status
                    .Cells(outputRow, 11).Value = rowData("Balance")
                    .Cells(outputRow, 11).NumberFormat = "#,##0"
                End With
                
                outputRow = outputRow + 1
                groupTotal = groupTotal + rowData("Balance")
            Next rowKey
        Next subGroupKey
        
        ' Add group total row to manager's sheet
        With outputSheet
            .Cells(outputRow, 3).Value = "Total"
            .Cells(outputRow, 11).Value = groupTotal
            .Range(.Cells(outputRow, 1), .Cells(outputRow, 11)).Interior.Color = RGB(240, 240, 240)
            .Range(.Cells(outputRow, 11), .Cells(outputRow, 11)).Font.Bold = True
            .Cells(outputRow, 11).NumberFormat = "#,##0"
        End With
        
        grandTotal = grandTotal + groupTotal
        
        ' Add manager's grand total to summary sheet
        summarySheet.Cells(summaryRow, 1).Value = projectManager
        summarySheet.Cells(summaryRow, 2).Value = groupTotal
        summaryRow = summaryRow + 1
    Next groupKey
    
    ' Add final grand total row in the summary sheet
    With summarySheet
        .Cells(summaryRow, 1).Value = "Overall Grand Total"
        .Cells(summaryRow, 2).Value = grandTotal
        .Range(.Cells(summaryRow, 1), .Cells(summaryRow, 2)).Font.Bold = True
    End With
    
    ' Auto-fit columns and add borders for all sheets
    For Each outputSheet In ThisWorkbook.Sheets
        If outputSheet.Name <> "Manager_Summary" Then
            outputSheet.UsedRange.Columns.AutoFit
            outputSheet.UsedRange.Borders.LineStyle = xlContinuous
        End If
    Next outputSheet

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Function WorksheetExists(sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function
