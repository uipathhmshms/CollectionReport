Sub Main(tableWidth As String)
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
    
    ' Get the last used row in the table (change this based on your table's range if needed)
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Apply background color (light blue) to the first, second, and third columns (excluding the first row)
    ApplyLightBlueToColumns lastRow
    
    ' Add a filter to the first row
    AddFilterToFirstRow firstRowRange
    
    ' Automatically adjust the width of each column to fit the content of the first row
    AutoFitColumns firstRowRange, intTableWidth
    
    ' Ensure no column exceeds a width of 20
    LimitColumnWidth intTableWidth, 20
End Sub

Sub ApplyFirstRowBackgroundColor(firstRowRange As Range)
    ' Apply the background color to the first row of the table
    firstRowRange.Interior.Color = RGB(16, 72, 97) ' RGB color for blue
End Sub

Sub ChangeFirstRowFontColor(firstRowRange As Range)
    ' Change the font color of the first row to white
    firstRowRange.Font.Color = RGB(255, 255, 255) ' RGB color for white
End Sub

Sub ApplyLightBlueToColumns(lastRow As Long)
    Dim firstColRange As Range
    Dim secondColRange As Range
    Dim thirdColRange As Range
    
    ' Apply background color (light blue) to the first, second, and third columns
    Set firstColRange = Range("A2:A" & lastRow) ' First column (A2 to the last row in column A)
    Set secondColRange = Range("B2:B" & lastRow) ' Second column (B2 to the last row in column B)
    Set thirdColRange = Range("C2:C" & lastRow) ' Third column (C2 to the last row in column C)
    
    firstColRange.Interior.Color = RGB(173, 216, 230) ' Light blue for the first column
    secondColRange.Interior.Color = RGB(173, 216, 230) ' Light blue for the second column
    thirdColRange.Interior.Color = RGB(173, 216, 230) ' Light blue for the third column
End Sub

Sub AddFilterToFirstRow(firstRowRange As Range)
    ' Add a filter to the first row (autofilter)
    firstRowRange.AutoFilter
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
