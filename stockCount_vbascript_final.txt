Sub stockCount():

'Assigning variables
Dim t As Double
Dim lastRow As Double
Dim openDate2018 As Double
Dim closeDate2018 As Double
Dim openDate2019 As Double
Dim closeDate2019 As Double
Dim openDate2020 As Double
Dim closeDate2020 As Double
Dim openPrice As Double
Dim closePrice As Double
Dim condition1 As FormatCondition
Dim condition2 As FormatCondition
Dim ws As Worksheet

'Extra variable for iteration
t = 1
'Variable to find the last row
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Assigning the values of the opening date and closing date
openDate2018 = 20180102
closeDate2018 = 20181231
openDate2019 = 20190102
closeDate2019 = 20191231
openDate2020 = 20200102
closeDate2020 = 20201231

'Loop through each worksheet
For Each ws In Worksheets

    'Set the headers to their respective names
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Delete any existing conditional formatting
    ws.Range("J:J").FormatConditions.Delete
    
    'Specify conditional formatting
    Set condition1 = ws.Range("J:J").FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set condition2 = ws.Range("J:J").FormatConditions.Add(xlCellValue, xlLess, "=0")
    
    'Set interior color to green if greater than 0
    With condition1
    .Interior.ColorIndex = 4
    End With
    
    'Set interior color to red if less than 0
    With condition2
    .Interior.ColorIndex = 3
    End With
    
    'Set the number format for column J, which holds the yearly change
    ws.Range("J:J").NumberFormat = "##0.00"

    'Set the number format for column K, which holds the percentage change
    ws.Range("K:K").NumberFormat = "##0.00%"

    'Loop through rows to find data and move to results columns
    For i = 2 To lastRow

        'FINDS TICKER, YEARLY DIFFERENCE, AND PERCENT DIFFERENCE
        'If the cell matches the opening date value, then
        If ws.Cells(i, 2).Value = openDate2020 Or ws.Cells(i, 2).Value = openDate2019 Or ws.Cells(i, 2).Value = openDate2018 Then
    
            'Log the ticker to the specified cell
            ws.Cells(t + 1, 9).Value = ws.Cells(i, 1).Value
        
            'Assign the opening price to a variable
            openPrice = ws.Cells(i, 3).Value
             
        'If the cell matches the closing date value, then
        ElseIf ws.Cells(i, 2).Value = closeDate2020 Or ws.Cells(i, 2).Value = closeDate2019 Or ws.Cells(i, 2).Value = closeDate2018 Then
    
            'Assign the closing price to a variable
            closePrice = ws.Cells(i, 6).Value
        
            'Log the difference between closing price and opening price
            ws.Cells(t, 10).Value = closePrice - openPrice
        
            'Log the relative change (percent change) between closing price and opening price
            ws.Cells(t, 11).Value = (closePrice - openPrice) / openPrice
        
        End If

        'FINDS TOTAL VOLUME
        'If the cell matches the cell before it, then
        If ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
    
            'Log the sums of the stock volume
            ws.Cells(t, 12).Value = ws.Cells(t, 12).Value + ws.Cells(i, 7).Value
        
        'If the cell does not match the cell before it, then
        Else
    
            'Increase the "t" variable, which moves onto the next row
            t = t + 1
        
            'Log the stock volume
            ws.Cells(t, 12).Value = ws.Cells(i, 7).Value
        
        End If
  
    'Next iteration
    Next i
    
'Reset value of 't', so each new worksheet will log values onto first row
t = 1

'Next worksheet
Next ws

End Sub