Sub AddStyleToSheet()
    Dim tableStart As Range
    Dim firstRowRange As Range
    Dim intTableWidth As Integer
    Dim lastRow As Long
    Dim usedRange As Range
    
    ' Get the used range in the worksheet
    Set usedRange = ActiveSheet.UsedRange
    
    ' Get the width (number of columns) of the used range
    intTableWidth = usedRange.Columns.Count
    
    ' Define the starting cell of the table (A1 in this case)
    Set tableStart = Range("A1")
    
    ' Set the range for the first row of the table based on the table's width (number of columns)
    Set firstRowRange = tableStart.Resize(1, intTableWidth)
    
    ' Apply background color to the first row
    ApplyFirstRowBackgroundColor firstRowRange
    
    ' Change the font color of the first row to white
    ChangeFirstRowFontColor firstRowRange
    
    ' Center the text in the first row
    CenterTextInFirstRow firstRowRange
    
    ' Get the last used row in the sheet 
    lastRow = usedRange.Rows.Count
    
    ' Change the background color of the first column (A)
    ChangeFirstColumnBackgroundColor lastRow
    
    ' Apply light blue background color to even index columns (B, D, F, etc.)
	ApplyLightBlueToEvenColumns lastRow, intTableWidth
    
    ' Add a filter to the first row
    AddFilterToFirstRow firstRowRange
    
    ' Freeze the first row
    FreezeFirstRow
    
    ' Automatically adjust the width of each column to fit the content of the first row
    AutoFitColumns firstRowRange, intTableWidth
    
    ' Ensure no column exceeds a width of 50
    'LimitColumnWidth intTableWidth, 50
    
    ' Center align all the text in the entire worksheet
    CenterAlignAllText
	
	' Change the background color of the "Status" and "יתרה" columns
    ChangeStatusAndBalanceColors lastRow
	
	SetSheetDirectionRTL
	
	FormatTotalRows
	
	MergeFirstColumnRowsExceptFirstAndLast
	
	SetSheetDirectionRTL
	
	ColorGrandTotalRow lastRow, intTableWidth
	
	FormatBigNumbers lastRow, intTableWidth
End Sub

sub mergefirstcolumnrowsexceptfirstandlast()
	dim ws as worksheet
	dim lastrow as long
	dim firstrow as long
	dim mergerange as range

	' ' ' set the worksheet to the active sheet (modify as needed)
	set ws = activesheet

	' ' ' define the range of rows to merge
	firstrow = 2 ' skip the first row (header)
	lastrow = ws.cells(ws.rows.count, 1).end(xlup).row ' get the last row with data in column a

	' ' ' ensure there are at least 3 rows to work with
	if lastrow <= firstrow then
		exit sub
	end if

	' define the range to merge
	set mergerange = ws.range(ws.cells(firstrow, 1), ws.cells(lastrow - 1, 1))

	' ' ' merge rows in the first column from the second row to the penultimate row
	mergerange.merge

	' ' ' align the text in the merged cell
	with mergerange
	.horizontalalignment = xlcenter ' center horizontally
	.verticalalignment = xltop ' align text to the top
	end with
end sub

Sub ChangeStatusAndBalanceColors(lastRow As Long)
    Dim statusCol As Integer
    Dim balanceCol As Integer
    Dim i As Long
    Dim statusCell As Range
    Dim balanceCell As Range
    
    ' Set the fixed column numbers for "Status" and "יתרה"
    statusCol = 10 ' Column number for "Status"
    balanceCol = 11 ' Column number for "יתרה"
    
    ' Loop through each row starting from row 2 to the last row
    For i = 2 To lastRow
        ' Get the status and balance cells in the current row
        Set statusCell = Cells(i, statusCol)
        Set balanceCell = Cells(i, balanceCol)
        
        ' Apply color logic based on the status value
        Select Case statusCell.Value
            Case "On Time"
                statusCell.Interior.Color = RGB(146, 208, 80) ' Green
                balanceCell.Interior.Color = RGB(146, 208, 80) ' Green
            Case "Delayed"
                statusCell.Interior.Color = RGB(255, 255, 0) ' Yellow
				balanceCell.Interior.Color = RGB(255, 255, 0) ' Yellow
            Case "Debt at Risk"
                statusCell.Interior.Color = RGB(247, 199, 172) ' Pink
                balanceCell.Interior.Color = RGB(247, 199, 172) ' Pink
        End Select
    Next i
End Sub

Sub ApplyFirstRowBackgroundColor(firstRowRange As Range)
    ' Apply the background color to the first row of the table
    firstRowRange.Interior.Color = RGB(16, 72, 97) ' RGB color for blue
End Sub

Sub ChangeFirstRowFontColor(firstRowRange As Range)
    ' Change the font color of the first row to white
    firstRowRange.Font.Color = RGB(255, 255, 255) ' RGB color for white
End Sub

Sub ChangeFirstColumnBackgroundColor(lastRow As Long)
    ' Change the background color of the first column (A) to RGB(68, 179, 225)
    Dim firstColRange As Range
    Set firstColRange = Range("A2:A" & (lastRow-1)) ' First column (A2 to last row in column-1 A (the last one is grand total))
    
    firstColRange.Interior.Color = RGB(68, 179, 225) ' Light blue color (RGB(68, 179, 225))
End Sub

Sub ApplyLightBlueToEvenColumns(lastRow As Long, intTableWidth As Integer)
    Dim col As Integer
    Dim row As Long
    Dim colRange As Range

    ' Loop through all even-indexed columns (B, D, F, etc.)
    For col = 2 To intTableWidth Step 2
        ' Loop through each cell in the column
        For row = 2 To lastRow
            ' Check if the cell does not already have a background color
            If Cells(row, col).Interior.ColorIndex = xlNone Then
                ' Apply light blue color
                Cells(row, col).Interior.Color = RGB(192, 230, 245)
            End If
        Next row
    Next col
End Sub

Sub AddFilterToFirstRow(firstRowRange As Range)
    ' Add a filter to the first row (autofilter)
    firstRowRange.AutoFilter
End Sub

Sub FreezeFirstRow()
    ' Freeze the first row
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub

Sub AutoFitColumns(firstRowRange As Range, intTableWidth As Integer)
    ' Automatically adjust the width of each column to fit the content of the first row
    firstRowRange.EntireColumn.AutoFit
End Sub

Sub LimitColumnWidth(intTableWidth As Integer, maxWidth As Integer)
    Dim col As Integer
    
    ' Loop through each column to make sure no column exceeds the max width
    For col = 1 To intTableWidth
        If Columns(col).ColumnWidth > maxWidth Then
            Columns(col).ColumnWidth = maxWidth
        End If
    Next col
End Sub

Sub CenterTextInFirstRow(firstRowRange As Range)
    ' Center the text in the first row
    firstRowRange.HorizontalAlignment = xlCenter
End Sub

Sub CenterAlignAllText()
    ' Get the used range in the active sheet
    Dim usedRange As Range
    Set usedRange = ActiveSheet.UsedRange
    
    ' Center align the text in the entire used range
    usedRange.HorizontalAlignment = xlCenter
End Sub

Sub FormatTotalRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
	Dim lastCol As Long

    ' Set the active sheet
    Set ws = ActiveSheet

    ' Get the last used row in column K
    lastRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
	lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Find last used column
	
    ' Loop through all rows in column G
    For i = 2 To lastRow
        ' Check if the month=0 - meaning thats a total row
        If ws.Cells(i, 2).Value="Total" Then
            ' Color the entire row blue
			ws.Range(ws.Cells(i, 2), ws.Cells(i, lastCol)).Interior.Color = RGB(131, 204, 235)
        End If
    Next i
End Sub

' Sub to set the sheet direction to Right-to-Left
Sub SetSheetDirectionRTL()
    Dim sheet As Worksheet
    
    ' Check if there are sheets in the workbook
    If ThisWorkbook.Sheets.Count = 0 Then
        MsgBox "No sheets found in the workbook!", vbExclamation
        Exit Sub
    End If
    
    ' Set the first sheet (or modify to target a specific sheet)
    Set sheet = ThisWorkbook.Sheets(1) ' Use Set for object assignment
    
    ' Apply Right-to-Left settings
    With sheet
        .DisplayRightToLeft = True ' Set sheet direction to Right-to-Left
        .Cells.HorizontalAlignment = xlHAlignRight ' Align text to the right (use xlHAlignRight for clarity)
    End With
End Sub

Sub ColorGrandTotalRow(lastRow As Long, intTableWidth As Integer)
    Dim ws As Worksheet
    Dim grandTotalRange As Range
    
    Set ws = ActiveSheet
    Set grandTotalRange = ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, intTableWidth))
    
    With grandTotalRange
        .Interior.Color = RGB(131, 204, 235) ' Dark blue color
        .Font.Bold = True ' Make text bold
    End With
End Sub

Sub FormatBigNumbers(lastRow As Long, intTableWidth As Integer)
    Dim ws As Worksheet
    Dim row As Long
    Dim col As Integer
    Dim cell As Range
    
    Set ws = ActiveSheet
    
    ' Loop through all cells in the used range (excluding header row)
    For row = 2 To lastRow
	col =11
	Set cell = ws.Cells(row, col)
	' Check if the cell contains a numeric value
	If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
		' Skip if cell contains "Total" or is in the Status column (column 10)
		If cell.Value <> "Total" And col <> 10 Then
			' Format numbers >= 1000 with thousand separators and 2 decimal places
			If Abs(cell.Value) >= 1000 Then
				cell.NumberFormat = "#,##0.00"
			Else
				cell.NumberFormat = "0.00"
			End If
		End If
	End If
    Next row
End Sub