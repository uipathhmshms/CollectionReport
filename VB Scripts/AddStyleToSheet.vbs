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
    LimitColumnWidth intTableWidth, 50
    
    ' Center align all the text in the entire worksheet
    CenterAlignAllText
	
	' Change the background color of the "Status" and "יתרה" columns
    ChangeStatusAndBalanceColors lastRow
	
	MergeFirstColumnRowsExceptFirstAndLast
	
End Sub

Sub MergeFirstColumnRowsExceptFirstAndLast()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim firstRow As Long
    Dim mergeRange As Range
    
    ' Set the worksheet to the active sheet (modify as needed)
    Set ws = ActiveSheet
    
    ' Define the range of rows to merge
    firstRow = 2 ' Skip the first row (header)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Get the last row with data in column A
    
    ' Ensure there are at least 3 rows to work with
    If lastRow <= firstRow Then
        MsgBox "Not enough rows to merge.", vbExclamation
        Exit Sub
    End If
    
    ' Define the range to merge
    Set mergeRange = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow - 1, 1))
    
    ' Merge rows in the first column from the second row to the penultimate row
    mergeRange.Merge
    
    ' Align the text in the merged cell
    With mergeRange
        .HorizontalAlignment = xlCenter ' Center horizontally
        .VerticalAlignment = xlTop ' Align text to the top
    End With
End Sub

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
            Case "On time"
                statusCell.Interior.Color = RGB(146, 208, 80) ' Green
                balanceCell.Interior.Color = RGB(146, 208, 80) ' Green
            Case "Delayed"
                statusCell.Interior.Color = RGB(255, 255, 0) ' Yellow
				balanceCell.Interior.Color = RGB(255, 255, 0) ' Yellow
            Case "Debt at risk"
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