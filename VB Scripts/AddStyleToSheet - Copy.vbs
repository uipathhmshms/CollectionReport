Sub AddStyleToSheet(tableWidth As String)
    Dim tableStart As Range
    Dim firstRowRange As Range
    Dim intTableWidth As Integer
    Dim lastRow As Long
    
    ' Convert the string argument to an integer
    intTableWidth = CInt(tableWidth)
    
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
    
    ' Get the last used row in the table 
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
	
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
    Set firstColRange = Range("A2:A" & lastRow) ' First column (A2 to last row in column A)
    
    firstColRange.Interior.Color = RGB(68, 179, 225) ' Light blue color (RGB(68, 179, 225))
End Sub

Sub ApplyLightBlueToEvenColumns(lastRow As Long, intTableWidth As Integer)
    Dim col As Integer
    Dim colRange As Range
    
    ' Loop through all columns and apply light blue to even indexed columns (B, D, F, etc.)
    For col = 2 To intTableWidth Step 2
        Set colRange = Range(Cells(2, col), Cells(lastRow, col)) ' Range from row 2 to last row in even columns
        colRange.Interior.Color = RGB(192, 230, 245) ' Light blue color
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