Sub GroupAndSumData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim resultRow As Long
    Dim managerName As String
    Dim sumBalance As Double
    
    ' Set the worksheet to the active sheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Replace "Sheet1" with your sheet's name
    
    ' Find the last row with data in column A (שם מנהל פרויקט)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Create a new worksheet for the result (optional)
    Set resultWs = ThisWorkbook.Sheets.Add
    resultWs.Name = "GroupedData"
    
    ' Set headers for the result sheet
    resultWs.Cells(1, 1).Value = "שם מנהל פרויקט"
    resultWs.Cells(1, 2).Value = "קוד חשבון"
    resultWs.Cells(1, 3).Value = "שם חשבון"
    resultWs.Cells(1, 4).Value = "סעיף תקציבי"
    resultWs.Cells(1, 5).Value = "אסמכתא"
    resultWs.Cells(1, 6).Value = "תאריך אסמכתא"
    resultWs.Cells(1, 7).Value = "תאריך לתשלום"
    resultWs.Cells(1, 8).Value = "תאריך פעילות"
    resultWs.Cells(1, 9).Value = "Status"
    resultWs.Cells(1, 10).Value = "Sum of יתרה"
    
    resultRow = 2 ' Start from row 2 for data
    
    ' Initialize variables
    managerName = ws.Cells(2, 1).Value ' First manager's name
    sumBalance = 0 ' Initialize sum of balance
    
    ' Loop through the data to group by manager and sum the balance
    For currentRow = 2 To lastRow
        ' Check if the manager name is the same as the previous one
        If ws.Cells(currentRow, 1).Value = managerName Then
            sumBalance = sumBalance + ws.Cells(currentRow, 10).Value ' Sum the balance (column J)
        Else
            ' If manager name changes, write the previous manager's data and reset the variables
            resultWs.Cells(resultRow, 1).Value = managerName
            resultWs.Cells(resultRow, 10).Value = sumBalance
            
            ' Move to next row in the result sheet
            resultRow = resultRow + 1
            
            ' Update the manager name and reset the sum
            managerName = ws.Cells(currentRow, 1).Value
            sumBalance = ws.Cells(currentRow, 10).Value ' Start summing for the new manager
        End If
        
        ' Copy the other details (column values) for the grouped rows
        resultWs.Cells(resultRow - 1, 2).Value = ws.Cells(currentRow, 2).Value ' קוד חשבון
        resultWs.Cells(resultRow - 1, 3).Value = ws.Cells(currentRow, 3).Value ' שם חשבון
        resultWs.Cells(resultRow - 1, 4).Value = ws.Cells(currentRow, 4).Value ' סעיף תקציבי
        resultWs.Cells(resultRow - 1, 5).Value = ws.Cells(currentRow, 6).Value ' אסמכתא
        resultWs.Cells(resultRow - 1, 6).Value = ws.Cells(currentRow, 7).Value ' תאריך אסמכתא
        resultWs.Cells(resultRow - 1, 7).Value = ws.Cells(currentRow, 8).Value ' תאריך לתשלום
        resultWs.Cells(resultRow - 1, 8).Value = ws.Cells(currentRow, 9).Value ' תאריך פעילות
        resultWs.Cells(resultRow - 1, 9).Value = "Status" ' You can define your status logic here
        
    Next currentRow
    
    ' Handle the last group of data after the loop ends
    resultWs.Cells(resultRow, 1).Value = managerName
    resultWs.Cells(resultRow, 10).Value = sumBalance
    
    ' Format the Sum of יתרה column as currency
    resultWs.Columns(10).NumberFormat = "#,##0.00"
    
    MsgBox "Grouping and Summing Completed!"
    
End Sub
