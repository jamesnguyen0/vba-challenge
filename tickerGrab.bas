Sub tickerGrab():

Dim ws As Worksheet
Dim stockEntries As Double

'input variables
Dim tickerTrailing As String
Dim tickerLeading As String
Dim ticker As String

'output variables
Dim priceArray(2) As Double 'store open price at 0 and close price at 1
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalVolume As Double

'output table variables
Dim row As Integer

'loop through each sheet in book
For Each ws In Worksheets

ws.Activate

'output table 1
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

stockEntries = WorksheetFunction.CountA(Range("A:A"))
row = 2 'start table at top of sheet, minus header

For i = 2 To stockEntries
       
    ticker = Cells(i, 1).Value
    tickerTrailing = Cells(i - 1, 1).Value
    tickerLeading = Cells(i + 1, 1).Value
        
    totalVolume = totalVolume + Cells(i, 7).Value
        
    'only execute and get open price if FIRST instance of ticker
    If Not (ticker = tickerTrailing) Then
        priceArray(0) = Cells(i, 3).Value
    End If
        
    'only execute if LAST instance of ticker
    If (Not (ticker = tickerLeading)) Then
        Cells(row, 9).Value = ticker
        Cells(row, 12).Value = totalVolume
        
        priceArray(1) = Cells(i, 6).Value
        yearlyChange = priceArray(1) - priceArray(0)
        
        'avoid divide by 0
        If (priceArray(0) = 0) Then
            percentChange = 0
        Else
            percentChange = yearlyChange / priceArray(0)
        End If
        
        Cells(row, 10).Value = yearlyChange
        
        'formatting
        If (yearlyChange > 0) Then
            Cells(row, 10).Interior.ColorIndex = 4
        ElseIf (yearlyChange < 0) Then
            Cells(row, 10).Interior.ColorIndex = 3
        End If
            
        Cells(row, 11).Value = percentChange
        
        'advance row in output table, clear holding variables
        row = row + 1
        
        priceArray(0) = 0
        priceArray(1) = 0
        yearlyChange = 0
        percentChange = 0
        totalVolume = 0
    End If
Next i

'table 1 formatting
Range("J:J").NumberFormat = "0.00"
Range("K:K").NumberFormat = "0.00%"
Range("I:L").Columns.AutoFit

'output table 2
Dim outputTableEntries As Double
outputTableEntries = WorksheetFunction.CountA(Range("I:I"))

Dim biggestIncrease As Double
Dim biggestDecrease As Double
Dim biggestVolume As Double

Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

For i = 2 To outputTableEntries
       
    If biggestIncrease < Cells(i, 11).Value Then
        biggestIncrease = Cells(i, 11).Value
        Cells(2, 15).Value = Cells(i, 9)
        Cells(2, 16).Value = biggestIncrease
        End If
        
    If biggestDecrease > Cells(i, 11).Value Then
        biggestDecrease = Cells(i, 11).Value
        Cells(3, 15).Value = Cells(i, 9)
        Cells(3, 16).Value = biggestDecrease
        End If
    
    If biggestVolume < Cells(i, 12).Value Then
        biggestVolume = Cells(i, 12).Value
        Cells(4, 15).Value = Cells(i, 9)
        Cells(4, 16).Value = biggestVolume
        End If
            
Next i

'reset holding variables for next loop
biggestIncrease = 0
biggestDecrease = 0
biggestVolume = 0

'table 2 formatting
Range("P2", "P3").NumberFormat = "0.00%"
Range("N:P").Columns.AutoFit

Next ws

End Sub





