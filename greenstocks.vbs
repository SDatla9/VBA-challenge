Sub Stockmarket_Analysis()
 
Dim total As Double
Dim start  As Double
Dim Ticker As String
Dim TickerOpenPrice, TickerClosedPrice As Double
 
For Each ws In Worksheets
 

    ' Set initial values
    total = 0
    Change = 0
    start = 2
    TickerOpenPrice = ws.Cells(2, "C").Value
    
    ' Set up titles for data table
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percentage Change"
    ws.Cells(1, "L").Value = "Total Stock Volume"

    ' get the row number of the last row with data
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
   

    For i = 2 To RowCount
        total = total + Cells(i, 7).Value
        Ticker = ws.Cells(i, 1).Value
     
        ' If ticker changes then print results
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(start, 9).Value = Ticker
        TickerClosedPrice = ws.Cells(i, 6).Value
        ws.Cells(start, 10).Value = TickerClosedPrice - TickerOpenPrice
        If TickerOpenPrice <> 0 Then
        ws.Cells(start, 11).Value = FormatPercent((TickerClosedPrice - TickerOpenPrice) / TickerOpenPrice, 2)
        Else
        ws.Cells(start, 11).Value = Null
        End If
        ws.Cells(start, 12).Value = total
        'Green for positive yearly change
        If ws.Cells(start, 10).Value > 0 Then
        ws.Cells(start, 10).Interior.ColorIndex = 4
        'Red for negative yearly change
        Else
        ws.Cells(start, 10).Interior.ColorIndex = 3
        End If

        
    TickerOpenPrice = ws.Cells(i + 1, 3).Value
    total = 0
    start = start + 1
    End If
    Next i
    
    
    ws.Columns("A:Q").AutoFit
    Next ws
        
      
End Sub