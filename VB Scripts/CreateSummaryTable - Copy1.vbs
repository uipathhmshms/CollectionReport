Sub CreateSummaryTable()
    Dim objExcel, objWorkbook, objSheet
	Dim rowTotal, rowSum, rowStatus
	Dim lastRow, i
	Dim grandTotal, totalDelayed, totalOnTime, totalDebtAtRisk, sumRemaining

	' Reference the active workbook and sheets
    Set objSheet = ThisWorkbook.Sheets(1) ' Input data is on the first sheet

	' Get the last row with data in the sheet
	lastRow = objSheet.Cells(objSheet.Rows.Count, 1).End(-4162).Row ' 4162 corresponds to xlUp

	' Calculate Grand Total (assuming it's in the last row, column 'Sum of יתרה')
	grandTotal = objSheet.Cells(lastRow, 11).Value ' Column 11 corresponds to "יתרה"

	' Initialize variables for totals
	totalDelayed = 0
	totalOnTime = 0
	totalDebtAtRisk = 0

	' Loop through the rows to calculate the totals based on the "Status" column (assumed to be column 10 for "Status")
	For i = 2 To lastRow ' Assuming data starts from row 2
		rowStatus = objSheet.Cells(i, 10).Value ' Column 10 is "Status"
		rowSum = objSheet.Cells(i, 11).Value ' Column 12 is "Sum of יתרה"
		
		If rowStatus = "Delayed" Then
			totalDelayed = totalDelayed + rowSum
		ElseIf rowStatus = "On time" Then
			totalOnTime = totalOnTime + rowSum
		ElseIf rowStatus = "Debt at risk" Then
			totalDebtAtRisk = totalDebtAtRisk + rowSum
		End If
	Next

	' Calculate percentages for the sub-table
	Dim percentDelayed, percentOnTime, percentDebtAtRisk
	percentDelayed = (totalDelayed / grandTotal) * 100
	percentOnTime = (totalOnTime / grandTotal) * 100
	percentDebtAtRisk = (totalDebtAtRisk / grandTotal) * 100

	' Create sub-table headers
	objSheet.Cells(lastRow + 2, 1).Value = "אחוז מסך החוב"
	objSheet.Cells(lastRow + 2, 2).Value = "Sum of יתרה"
	objSheet.Cells(lastRow + 2, 3).Value = "Status"

	' Fill the sub-table with calculated values
	objSheet.Cells(lastRow + 3, 1).Value = percentDelayed & "%" ' Percentage for Delayed
	objSheet.Cells(lastRow + 3, 2).Value = totalDelayed
	objSheet.Cells(lastRow + 3, 3).Value = "Delayed"

	objSheet.Cells(lastRow + 4, 1).Value = percentOnTime & "%" ' Percentage for On time
	objSheet.Cells(lastRow + 4, 2).Value = totalOnTime
	objSheet.Cells(lastRow + 4, 3).Value = "On time"

	objSheet.Cells(lastRow + 5, 1).Value = percentDebtAtRisk & "%" ' Percentage for Debt at risk
	objSheet.Cells(lastRow + 5, 2).Value = totalDebtAtRisk
	objSheet.Cells(lastRow + 5, 3).Value = "Debt at risk"

	' Add the total row for 100% (grand total)
	objSheet.Cells(lastRow + 6, 1).Value = "100%"
	objSheet.Cells(lastRow + 6, 2).Value = grandTotal
	objSheet.Cells(lastRow + 6, 3).Value = "grandTotal"
	
	' Call the separate subroutine to apply styling
    Call ApplySummaryTableStyling(objSheet, lastRow)

End Sub

' Separate subroutine to apply styling to the summary table
Sub ApplySummaryTableStyling(objSheet, lastRow)
	objSheet.Cells(lastRow + 3, 1).Font.Name = "David"  ' Set the font to a Hebrew-compatible font (e.g., "David")


    ' Apply background color based on the Status
    ' "On time" - Green
    ' "Delayed" - Yellow
    ' "Debt at risk" - Pink

    ' Apply styling for "On time"
    objSheet.Range(objSheet.Cells(lastRow + 4, 1), objSheet.Cells(lastRow + 4, 3)).Interior.Color = RGB(146, 208, 80) ' Green

    ' Apply styling for "Delayed"
    objSheet.Range(objSheet.Cells(lastRow + 3, 1), objSheet.Cells(lastRow + 3, 3)).Interior.Color = RGB(255, 255, 0) ' Yellow

    ' Apply styling for "Debt at risk"
    objSheet.Range(objSheet.Cells(lastRow + 5, 1), objSheet.Cells(lastRow + 5, 3)).Interior.Color = RGB(247, 199, 172) ' Pink

    ' Apply styling for the "grandTotal" row 
    objSheet.Range(objSheet.Cells(lastRow + 6, 1), objSheet.Cells(lastRow + 6, 3)).Interior.Color = RGB(131, 204, 235) ' (for grandTotal row)
End Sub