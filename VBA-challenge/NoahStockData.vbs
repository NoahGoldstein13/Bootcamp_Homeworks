Sub NoahStockData()
'Remaining Tasks: make code pretty

'set variable for worksheets and start cycle through all worksheets in workbook
Dim ws As Worksheet
For Each ws In Worksheets

'set variable for holding ticker symbols
Dim ticker As String

'set variable for holding a stocks last close price
Dim closeprice As Double

'set variable for holding a stock's first open price and set the first variable equal to the first stock's open price
Dim openprice As Double
openprice = ws.Range("c2").Value

'set varibale for holding a stock's yearly price change
Dim yearlychange As Double

'set variable for holding a stock's yearly percent price change
Dim percentchange As Double

'set variable for holding a stock's total volume for a given year
Dim TotalVolume As Double

'set variable for the starting row of the summary table and set the first variable equal to 2
Dim SummaryTableRow As Integer
SummaryTableRow = 2

'Loop through all rows in the first column of each worksheet
For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Searches for when the value of the next cell is different than that of the current cell in the loop
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'sets value for the ticker variable and directs where each should be placed in the summary table
    ticker = ws.Cells(i, 1).Value
    ws.Range("I" & SummaryTableRow).Value = ticker
    
    'sets value for the closeprice and yearly change variables and instructs where to put yearlychange in the summary table
    closeprice = ws.Cells(i, 6).Value
    yearlychange = closeprice - openprice
    ws.Range("J" & SummaryTableRow).Value = yearlychange
    
    'avoids div by 0 error when openprice is equal to 0 in percentchange equation
    If openprice = 0 Then
    percentchange = 0
    Else
    percentchange = (closeprice - openprice) / openprice
    End If
    
    'puts percentchange into the proper place in the summary table
    ws.Range("K" & SummaryTableRow).Value = percentchange
    
    'adds the last day's volume for a ticker to the running total and puts total volume in the summary table
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    ws.Range("L" & SummaryTableRow).Value = TotalVolume
    
    'grabs new open price
    openprice = ws.Cells(i + 1, 3).Value
    
    'Next summary table row
    SummaryTableRow = SummaryTableRow + 1

    'Reset Total Volume
    TotalVolume = 0
    
    'continuously adds to a ticker's total volume
    Else
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value

    End If

Next i

'updates summary table headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'updates format for percent change
ws.Columns("K").NumberFormat = "0.00%"

'creates loop for updating yearly change colors
For i = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row

If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
Else
ws.Cells(i, 10).Interior.ColorIndex = 3
End If

Next i

Next ws

End Sub
