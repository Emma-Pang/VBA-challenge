Sub stockmarket():

'speed up 
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

'change in EACH WS
For Each ws In Worksheets

    'define ticker symbol
    Dim tickersymbol As String

    'keep track of new info in summary table
    Dim summarytablerow As Integer
    summarytablerow = 1
    'define opening and closing values

    Dim tickervolumetotal As Double
    tickervolumetotal = 0

    'define closing and opening price
    Dim closingvalue As Double
    Dim openingvalue As Double

    'define yearly change
    Dim yearlychange As Double
    yearlychange = 0

    Dim percentchange As Double
    percentchange = 0

    'determine last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'loop through rows
    
    For i = 1 To LastRow
        'get close price
        closingvalue = ws.Cells(i + 1, 6).Value
        
        'check if still within same ticker symbol and if it is not
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'set ticker name
        tickersymbol = ws.Cells(i, 1).Value

        'get opening value
        openingvalue = ws.Cells(i + 1, 3).Value

        'print tickersymbol in summary table
        ws.Range("I" & summarytablerow).Value = tickersymbol

        'print total volume per ticker to summary table
        ws.Range("L" & summarytablerow).Value = tickervolumetotal

        'print yearly change
        ws.Range("J" & summarytablerow).Value = yearlychange

        'print yearly change
        ws.Range("K" & summarytablerow).Value = percentchange
            
            
            'add one to summary table row
            summarytablerow = summarytablerow + 1
            tickervolumetotal = 0
            yearlychange = 0
            percentchange = 0

            
            'if cell immediately following a row is the same brand
            Else
            tickervolumetotal = tickervolumetotal + ws.Cells(i, 7).Value

            yearlychange = closingvalue - openingvalue

            If openingvalue > 0 Then
            percentchange = yearlychange / openingvalue
            Else
            percentchange = yearlychange / 1
            End If
        
        End If
        
    Next i



'percent change from opening price at the begining to closing price
'total volume of stock
    'label summary table rows
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    ws.Columns("I:L").AutoFit

    For i = 2 To LastRow

        ws.Cells(i, 11).NumberFormat = "0.00%"

        If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i

    

Next ws

Application.Calculation = xlCalculationAutomatic

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub