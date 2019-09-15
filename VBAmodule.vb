Sub TickerVolume()

'Set variables
Dim Ticker_Symbol As String
Dim Ticker_Volume As Double

'Determine the last row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Set the initial value for the ticker volume to 0
Ticker_Volume = 0

'Set location for each ticker symbol and total ticker volume
Dim Summary_Ticker_Row As Long
Summary_Ticker_Row = 2

'Loop through all rows
For i = 2 To LastRow

    'Check if the ticker symbol is still the same
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Set the ticker symbol
    Ticker_Symbol = Cells(i, 1).Value
    
    'Add to the ticker volume
    Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
    
    'Print the ticker symbol under new ticker column
    Range("I" & Summary_Ticker_Row).Value = Ticker_Symbol
    
    'Print the ticker volume under the total stock volume column
    Range("J" & Summary_Ticker_Row).Value = Ticker_Volume
    
    'Add one to the summary ticker row
    Summary_Ticker_Row = Summary_Ticker_Row + 1
    
    'Reset the ticker volume
    Ticker_Volume = 0
    
    'If the cell immediatelly following a row is the same ticker symbol
    Else
    
    'Add the ticker volume
    Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
    
    End If
    
    Next i
    
End Sub
