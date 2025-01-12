Option Explicit

Sub FormatLimitedTable()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Initialize variables
    Dim inputSheet As Worksheet
    Dim outputSheet As Worksheet
    Dim inputLastRow As Long
    Dim groupedData As Object
    Dim grandTotal As Double
    Dim today As Date
    Dim outputRow As Long
    
    ' Check if workbook has at least one sheet
    If ThisWorkbook.Sheets.Count < 1 Then
        MsgBox "Error: Workbook must contain at least one sheet.", vbCritical
        Exit Sub
    End If
    
    ' Reference the active workbook and sheets
    Set inputSheet = ThisWorkbook.Sheets("Limited")
    
    ' Check if input sheet has data
    If Application.WorksheetFunction.CountA(inputSheet.Cells) = 0 Then
        MsgBox "Error: Input sheet is empty.", vbCritical
        Exit Sub
    End If
    
    ' Create output sheet
    Set outputSheet = CreateOutputSheet(inputSheet)
    
    ' Get the last row in the input sheet
    inputLastRow = inputSheet.Cells(inputSheet.Rows.Count, 1).End(xlUp).Row
    If inputLastRow < 2 Then
        MsgBox "Error: No data found in input sheet.", vbCritical
        Exit Sub
    End If
    
    ' Initialize grouped data structure and grand total
    Set groupedData = CreateObject("Scripting.Dictionary")
    grandTotal = 0
    outputRow = 2 ' Start writing from the second row
    today = Date
    
    ' Process data and group it
    ProcessInputData inputSheet, inputLastRow, groupedData
    
    ' Write data to the output sheet
    WriteGroupData outputSheet, groupedData, outputRow, grandTotal, today
    
    ' Write grand total to the output sheet
    WriteGrandTotal outputSheet, grandTotal, outputRow
    
    ' Format the output sheet
    FormatSheet outputSheet    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' Function to create and format the output sheet
Function CreateOutputSheet(inputSheet As Worksheet) As Worksheet
    Dim outputSheet As Worksheet
    Dim sheetName As String
    sheetName = "Limited Collection Report"
    
    ' Create new sheet with unique name
    Set outputSheet = ThisWorkbook.Sheets.Add
    outputSheet.Name = sheetName
    
    ' Copy headers from input sheet
    inputSheet.Rows(1).Copy
    outputSheet.Rows(1).PasteSpecial xlPasteValues
    outputSheet.Rows(1).PasteSpecial xlPasteFormats

    ' Add the Status column header
    Dim lastColumn As Long
    lastColumn = outputSheet.Cells(1, outputSheet.Columns.Count).End(xlToLeft).Column
	outputSheet.Cells(1, lastColumn+1).Value = outputSheet.Cells(1, lastColumn).Value ' Push 'sum of yetra' one column to the right
    outputSheet.Cells(1, lastColumn).Value = "Status"
    
    ' Format the header row
    With outputSheet.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
    End With
    
    Application.CutCopyMode = False  ' Clear the clipboard
    Set CreateOutputSheet = outputSheet
End Function

' Function to process the input data and group it
Sub ProcessInputData(inputSheet As Worksheet, inputLastRow As Long, ByRef groupedData As Object)
    Dim dataRange As Range
    Dim dataArr As Variant
    Dim inputRow As Long
    Dim projectManager As String, accountCode As String, accountName As String
    Dim budgetSection As String, budgetSectionExtra As String
    Dim reference As String, referenceDate As Date
    Dim paymentDate As Date, activityDate As Date
    Dim balance As Double
    Dim groupKey As Variant, subGroupKey As Variant
    Dim rowData As Object

    ' Read data into array for faster processing
    Set dataRange = inputSheet.Range("A2:J" & inputLastRow)
    dataArr = dataRange.Value
    
    ' Process each row of data
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
        On Error GoTo 0
        
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
End Sub

' Function to write grouped data to the output sheet
Sub WriteGroupData(outputSheet As Worksheet, groupedData As Object, ByRef outputRow As Long, ByRef grandTotal As Double, today As Date)
    Dim groupKey As Variant, subGroupKey As Variant, rowKey As Variant
    Dim rowData As Object
    Dim groupTotal As Double
    Dim status As String, lastGroupKey As String
    lastGroupKey = ""
    
    ' Process each group and subgroup
    For Each groupKey In groupedData.Keys
        groupTotal = 0
        Dim splitGroup() As String
        splitGroup = Split(CStr(groupKey), "|")
        
        For Each subGroupKey In groupedData(groupKey).Keys
            For Each rowKey In groupedData(groupKey)(subGroupKey).Keys
                Set rowData = groupedData(groupKey)(subGroupKey)(rowKey)
                
                ' Determine status
                status = GetStatus(rowData("PaymentDueDate"), today)
                
                ' Write row to the output sheet
                WriteRowToOutputSheet outputSheet, rowData, status, outputRow, groupKey, lastGroupKey, subGroupKey
                
                groupTotal = groupTotal + rowData("Balance")
            Next rowKey
        Next subGroupKey
        
        ' Write the group total row
        WriteGroupTotal outputSheet, groupTotal, outputRow
        grandTotal = grandTotal + groupTotal
    Next groupKey
End Sub

' Function to determine the status based on the payment date
Function GetStatus(paymentDate As Variant, today As Date) As String
    If IsDate(paymentDate) Then
        Dim daysDelayed As Long
        daysDelayed = DateDiff("d", paymentDate, today)
        
        If daysDelayed < 0 Then
            GetStatus = "On time" ' תאריך הערך / התשלום טרם הגיע.
        ElseIf daysDelayed <= 45 Then
            GetStatus = "Delayed" ' תאריך הערך / התשלום בעיכוב של עד 45 יום.
        Else
            GetStatus = "Debt at risk" ' תאריך הערך / התשלום הינו בעיכוב של מעל 45 יום.
        End If
    Else
        GetStatus = "Invalid date"
    End If
End Function

' Function to write a row to the output sheet
Sub WriteRowToOutputSheet(outputSheet As Worksheet, rowData As Object, status As String, ByRef outputRow As Long, groupKey As Variant, ByRef lastGroupKey As String, subGroupKey As Variant)
    With outputSheet
        If CStr(groupKey) <> lastGroupKey Then
            .Cells(outputRow, 1).Value = Split(CStr(groupKey), "|")(0)
            .Cells(outputRow, 2).Value = Split(CStr(groupKey), "|")(1)
            .Cells(outputRow, 3).Value = rowData("AccountName")
            lastGroupKey = CStr(groupKey)
        Else
            .Cells(outputRow, 1).Value = ""
            .Cells(outputRow, 2).Value = ""
            .Cells(outputRow, 3).Value = ""
        End If
        
        .Cells(outputRow, 4).Value = Split(subGroupKey, "|")(1)
        .Cells(outputRow, 5).Value = rowData("BudgetSectionExtra")
        .Cells(outputRow, 6).Value = rowData("Reference")
        .Cells(outputRow, 7).Value = rowData("ReferenceDate")
        .Cells(outputRow, 8).Value = rowData("PaymentDueDate")
        .Cells(outputRow, 9).Value = rowData("ActivityDate")
        .Cells(outputRow, 10).Value = status
        .Cells(outputRow, 11).Value = rowData("Balance")
        
        ' Number formatting for the balance
        .Cells(outputRow, 11).NumberFormat = "#,##0"
    End With
    outputRow = outputRow + 1
End Sub

' Function to write the group total row
Sub WriteGroupTotal(outputSheet As Worksheet, groupTotal As Double, ByRef outputRow As Long)
    With outputSheet
        .Cells(outputRow, 3).Value = "Total"
        .Cells(outputRow, 11).Value = groupTotal
        .Range(.Cells(outputRow, 1), .Cells(outputRow, 11)).Interior.Color = RGB(131, 204, 235)
        .Range(.Cells(outputRow, 11), .Cells(outputRow, 11)).Font.Bold = True
        .Cells(outputRow, 11).NumberFormat = "#,##0"
    End With
    outputRow = outputRow + 1
End Sub

' Function to write the grand total row
Sub WriteGrandTotal(outputSheet As Worksheet, grandTotal As Double, ByRef outputRow As Long)
    With outputSheet
        .Cells(outputRow, 1).Value = "Grand Total"
        .Cells(outputRow, 11).Value = grandTotal
        .Range(.Cells(outputRow, 1), .Cells(outputRow, 11)).Interior.Color = RGB(131, 204, 235)
        .Range(.Cells(outputRow, 1), .Cells(outputRow, 11)).Font.Bold = True
        .Cells(outputRow, 11).NumberFormat = "#,##0"
    End With
    outputRow = outputRow + 1
End Sub

' Function to format the output sheet
Sub FormatSheet(outputSheet As Worksheet)
    With outputSheet
        .UsedRange.Columns.AutoFit
        .UsedRange.Borders.LineStyle = xlContinuous
    End With
End Sub
