Attribute VB_Name = "Module1"


' VBA Homework - The VBA of Wall Street

' Solution # 1


Sub tickerName()

    
'Set a variable for ticker name and column
        Dim tickerName As String
        
'Set a variable for total count of the stock total volume
        Dim Stockvolume As Double
        Stockvolume = 0

'Log each ticker name in the summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
'Label the Summary Table headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Total Stock Volume"

'Count the number of rows in the first column.
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through the rows by the ticker names

            
For i = 2 To lastrow

'Action for value of the next cell if different than of the current cell
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
'Set the ticker name
                tickerName = Cells(i, 1).Value
                
'Add the volume of stock
                Stockvolume = Stockvolume + Cells(i, 7).Value

'Print the ticker name in the summary table
                Range("I" & summary_ticker_row).Value = tickerName
                
'Print the trade volume for each ticker in the summary table
                Range("J" & summary_ticker_row).Value = Stockvolume

'Add +1 to the summary_ticker_row
                summary_ticker_row = summary_ticker_row + 1
                
'Reset stockvolume to zero
                Stockvolume = 0

            Else
              
'Add the volume of stock
          Stockvolume = Stockvolume + Cells(i, 7).Value

            End If
        
        Next i

End Sub
