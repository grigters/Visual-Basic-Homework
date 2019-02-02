Sub StockMarket()

'create variables
Dim total As Double
Dim next_ticker As Integer
Dim ticker As String
Dim ws As Worksheet
Dim openprice As Double
Dim closeprice As Double
Dim change As Double
Dim lastrow As Double

'add values to created variables
total = 0
next_ticker = 2

'start the for each loop
For Each ws In Worksheets

'add the open stock price to variable
openprice = ws.Cells(2, 3).Value

For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
    total = total + ws.Cells(i, 7).Value
Else
    ticker = ws.Cells(i, 1).Value
    total = total + ws.Cells(i, 7).Value
    ws.Cells(next_ticker, 9).Value = ticker
    ws.Cells(next_ticker, 12).Value = total
    total = 0
    
    'add the close stock price to variable
    closeprice = ws.Cells(i, 6).Value
    
    'change of stock price at year end compared to beginning + cell formatting
    ws.Cells(next_ticker, 10).Value = closeprice - openprice
    ws.Cells(next_ticker, 11).NumberFormat = "00.000000000"
    change = ws.Cells(next_ticker, 10).Value
    
    'if statement to change the color depending on if the change is positive or negative
    If change > 0 Then
    ws.Cells(next_ticker, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(next_ticker, 10).Interior.ColorIndex = 3
    End If
    
    'percent change with cell formatting + if statement for PLNT stock with 0 change
    If change = 0 Then
    ws.Cells(next_ticker, 11) = 0
    'can't divide by 0 so leaving the cell blank
    ElseIf openprice = 0 Then
    ws.Cells(next_ticker, 11).Value = ""
    Else
    ws.Cells(next_ticker, 11).Value = change / openprice
    ws.Cells(next_ticker, 11).NumberFormat = "0.00%"
    End If
    
    'set the open price for the next stock
    openprice = ws.Cells(i + 1, 3).Value
    
    next_ticker = next_ticker + 1
    
End If
Next i
    
    'add header text to columns and rows
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    '--------------Hard Assignment ---------------
    'add headers
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Greatest Increase
    max_num1 = WorksheetFunction.Max(ws.Columns(11))
    max_stock1 = WorksheetFunction.Match(max_num1, ws.Columns(11), 0)
    ws.Cells(2, 16).Value = ws.Cells(max_stock1 + 1, 9)
    ws.Cells(2, 17).Value = max_num1
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    'Greatest Decrease
    min_num1 = WorksheetFunction.Min(ws.Columns(11))
    min_stock1 = WorksheetFunction.Match(min_num1, ws.Columns(11), 0)
    ws.Cells(3, 16).Value = ws.Cells(min_stock1 + 1, 9)
    ws.Cells(3, 17).Value = min_num1
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    'Greatest Volume Increase
    max_num1 = WorksheetFunction.Max(ws.Columns(12))
    max_stock1 = WorksheetFunction.Match(max_num1, ws.Columns(12), 0)
    ws.Cells(4, 16).Value = ws.Cells(max_stock1 + 1, 9)
    ws.Cells(4, 17).Value = max_num1
    
    'Autofit to display data
    ws.Columns("A:Q").AutoFit
    
    'reset
    next_ticker = 2
Next ws


End Sub