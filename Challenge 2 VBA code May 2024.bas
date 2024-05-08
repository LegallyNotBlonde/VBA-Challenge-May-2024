Attribute VB_Name = "Module1"
Sub StockChange()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Total_Stock_Volume As Currency ' Change data type to Currency
    Dim Open_Price As Double
    Dim Quarterly_Change As Double
    Dim Percentage_Change As Double
    Dim i As Long
    Dim b As Long
    Dim MaxIncrTicker As String
    Dim MaxDecrTicker As String
    Dim MaxVolTicker As String
    Dim Max_Volume As Double
    Dim Gr_Incr As Double
    Dim Gr_Decr As Double
    
    For Each ws In Worksheets 'to enter volue on all tabs
        'insert text in headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Total_Stock_Volume = 0
        b = 2 ' starting row for data
        
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            Open_Price = ws.Cells(i, 3).Value
            Total_Stock_Volume = CDbl(ws.Cells(i, 7).Value) + Total_Stock_Volume ' Convert to Double before adding
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(b, 9).Value = ws.Cells(i, 1).Value ' Ticker
                ws.Cells(b, 10).Value = ws.Cells(i, 6).Value - Open_Price ' Quarterly Change
                ws.Cells(b, 11).Value = (ws.Cells(i, 6).Value - Open_Price) / Open_Price ' Percentage Change
                ws.Cells(b, 12).Value = Total_Stock_Volume ' Total Stock Volume
                b = b + 1 ' increment row index
                Total_Stock_Volume = 0 ' reset Total_Stock_Volume
            End If
        Next i
        
        'determine max stock value, max % increase, and max % decrease
        Max_Volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
        ws.Cells(4, 17) = Max_Volume
        Gr_Incr = Application.WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(2, 17) = Gr_Incr
        Gr_Decr = Application.WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(3, 17) = Gr_Decr
        
        'Set the value for percentage change displayed in a percentage format
        ws.Range("K2:K5000").NumberFormat = "0.00%"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0"
        ws.Range("L2:L10000").NumberFormat = "0"
        
        'use index and match function to enter ticker for max increase, decrease, and the greatest stick volume
        MaxIncrTicker = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(Gr_Incr, ws.Range("K:K"), 0))
        MaxDecrTicker = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(Gr_Decr, ws.Range("K:K"), 0))
        MaxVolTicker = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(Max_Volume, ws.Range("L:L"), 0))
        
        ws.Range("P2") = MaxIncrTicker
        ws.Range("P3") = MaxDecrTicker
        ws.Range("P4") = MaxVolTicker
    Next ws
End Sub

Sub Formatting_Colors()
    Dim ws As Worksheet
    Dim Col As Integer
    Dim i As Long ' Declared var. as long basause of many rows

    ' Loop through each worksheet in the workbook
    For Each ws In Worksheets
        Col = 11 ' Applying colors to column K (11th column)

        ' Loop through rows in the current worksheet
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
            ' Set colors based on values (green for >0, red for <0, and default for 0)
            If ws.Cells(i, Col).Value > 0 Then
                ws.Cells(i, Col).Interior.ColorIndex = 4 ' Green
            ElseIf ws.Cells(i, Col).Value < 0 Then
                ws.Cells(i, Col).Interior.ColorIndex = 3 ' Red
            Else
                ws.Cells(i, Col).Interior.ColorIndex = 2 ' Default color
            End If
        Next i
    Next ws
End Sub

