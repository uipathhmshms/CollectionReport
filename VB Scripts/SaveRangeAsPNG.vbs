Sub SaveRangeAsPNG()
    Dim rng As Range
    Dim chart As ChartObject
    Dim ws As Worksheet

    ' Specify the sheet where your table is located (you can modify the sheet name here)
    Set ws = ThisWorkbook.Sheets("Summary") ' Replace "Sheet1" with your sheet's actual name
    
    ' Set the range you want to export 
    Set rng = ws.Range("A1:C5") 
    
    ' Create a temporary chart
    Set chart = ws.ChartObjects.Add(Left:=rng.Left, Top:=rng.Top, Width:=rng.Width, Height:=rng.Height)
    'chart.Chart.ChartArea.Clear ' Clear the chart area to avoid a chart background
    
    ' Copy the range as a picture and paste it into the chart
    rng.CopyPicture xlScreen, xlPicture
    chart.Chart.Paste

    ' Save the chart as PNG
    chart.Chart.Export Filename:="C:\Users\GiusRpa\Downloads\lidor\TableImage.png", FilterName:="PNG"

    ' Delete the temporary chart
    chart.Delete
End Sub
