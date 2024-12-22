Option Explicit

Sub TransformData()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim inputSheet As Worksheet
    Dim outputSheet As Worksheet
    Dim inputLastRow As Long
    Dim outputRow As Long
    Dim groupedData As Object
    Dim grandTotal As Double
    Dim today As Date
    Dim paymentDate As Date
    Dim status As String
    Dim groupKey As Variant
    Dim subGroupKey As Variant
    Dim splitGroup As Variant
    Dim rowKey As Variant
    Dim rowData As Object
    Dim groupTotal As Double
    Dim lastGroupKey As String
    
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
    
    ' Create new sheet with unique name
    Dim sheetName As String
    'sheetName = "Transformed_Data_" & Format(Now, "yyyymmdd_hhmmss")
	sheetName = "Transformed_Data_"
    Set outputSheet = ThisWorkbook.Sheets.Add
    outputSheet.Name = sheetName
    
    ' Copy headers from input sheet
	inputSheet.Rows(1).Copy
	outputSheet.Rows(1).PasteSpecial xlPasteValues
	outputSheet.Rows(1).PasteSpecial xlPasteFormats

	' Add the Status column header
	outputSheet.Cells(1, outputSheet.Cells(1, Columns.Count).End(xlToLeft).Column + 1).Value = "Status"

	' Format the header row
	With outputSheet.Rows(1)
		.Font.Bold = True
		.Interior.Color = RGB(220, 220, 220)
	End With

Application.CutCopyMode = False  ' Clear the clipboard
    
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
    
    ' Group data by "שם מנהל פרויקט" and "קוד חשבון"
    Dim inputRow As Long
    Dim dataRange As Range
    Set dataRange = inputSheet.Range("A2:J" & inputLastRow)
    
    ' Read data into array for faster processing
    Dim dataArr As Variant
    dataArr = dataRange.Value
    
    For inputRow = 1 To UBound(dataArr)
        Dim projectManager As String, accountCode As String, accountName As String
        Dim budgetSection As String, budgetSectionExtra As String
        Dim reference As String, referenceDate As Date, activityDate As Date
        Dim balance As Double
        
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
        
        groupKey = projectManager & "|" & accountCode
        
        If Not groupedData.Exists(groupKey) Then
            groupedData.Add groupKey, CreateObject("Scripting.Dictionary")
        End If
        
        subGroupKey = accountName & "|" & budgetSection
        If Not groupedData(groupKey).Exists(subGroupKey) Then
            groupedData(groupKey).Add subGroupKey, CreateObject("Scripting.Dictionary")
        End If
        
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
        splitGroup = Split(CStr(groupKey), "|")
        projectManager = splitGroup(0)
        accountCode = splitGroup(1)
        
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
                
                ' Write row to the output file
                With outputSheet
                    ' Only write project manager, account code, and account name if it's a new group
                    If CStr(groupKey) <> lastGroupKey Then
                        .Cells(outputRow, 1).Value = projectManager
                        .Cells(outputRow, 2).Value = accountCode
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
                    
                    ' Add number formatting for the balance
                    .Cells(outputRow, 11).NumberFormat = "#,##0"
                End With
                
                outputRow = outputRow + 1
                groupTotal = groupTotal + rowData("Balance")
            Next rowKey
        Next subGroupKey
        
        ' Add group total row with formatting
        With outputSheet
            .Cells(outputRow, 3).Value = rowData("AccountName") & " Total"
            .Cells(outputRow, 11).Value = groupTotal
            .Range(.Cells(outputRow, 1), .Cells(outputRow, 11)).Interior.Color = RGB(240, 240, 240)
            .Range(.Cells(outputRow, 11), .Cells(outputRow, 11)).Font.Bold = True
            .Cells(outputRow, 11).NumberFormat = "#,##0"
        End With
        
        outputRow = outputRow + 1
        grandTotal = grandTotal + groupTotal
        lastGroupKey = "" ' Reset for next group
    Next groupKey
    
    ' Add grand total row with formatting
    With outputSheet
        .Cells(outputRow, 1).Value = "Grand Total"
        .Cells(outputRow, 11).Value = grandTotal
        .Range(.Cells(outputRow, 1), .Cells(outputRow, 11)).Interior.Color = RGB(200, 200, 200)
        .Range(.Cells(outputRow, 1), .Cells(outputRow, 11)).Font.Bold = True
        .Cells(outputRow, 11).NumberFormat = "#,##0"
    End With
    
    ' Auto-fit columns and add borders
    With outputSheet
        .UsedRange.Columns.AutoFit
        .UsedRange.Borders.LineStyle = xlContinuous
    End With
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Transformation complete. Data has been written to sheet '" & sheetName & "'.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' Helper function to check if sheet exists [remains the same]