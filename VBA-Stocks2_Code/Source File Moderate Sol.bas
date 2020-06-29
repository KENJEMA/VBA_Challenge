Attribute VB_Name = "Module11"

' Solution # 2 (Moderate)


Sub tickerSymbol()

    
'Set a variable for ticker name and column
        Dim tickername As String
        
'Set a variable for total count of the stock total volume
        Dim Stockvolume As Double
        Stockvolume = 0

'Log each ticker name in the summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        
' Declare Yearly change as difference in closing price of the year and opening price
' Percent change is therefor, ((closing price - opening price)/opening Price)*100
        
        Dim Open_price As Double
        Open_price = Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim Percent_change As Double
        
        
        
'Label the Summary Table headers

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

'Count the number of rows in the first column.
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through the rows by the ticker names

            
For i = 2 To lastrow

'Action for value of the next cell if different than of the current cell
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
'Set the ticker name
                tickername = Cells(i, 1).Value
                
'Add the volume of stock
                Stockvolume = Stockvolume + Cells(i, 7).Value

'Print the ticker name in the summary table
                Range("I" & summary_ticker_row).Value = tickername
                
'Print the trade volume for each ticker in the summary table
                Range("J" & summary_ticker_row).Value = Stockvolume
                
                
'Define Close_price
                close_price = Cells(i, 6).Value

'Calculate yearly change
              yearly_change = (close_price - Open_price)

'Print the yearly change for each ticker in the summary table
              Range("J" & summary_ticker_row).Value = yearly_change

'Define parameters for percent change
                
                If (Open_price = 0) Then

                    Percent_change = 0

                Else
                    
                    Percent_change = yearly_change / Open_price
                
                End If
                
'Print the yearly change for each ticker in the summary table
              
              Range("K" & summary_ticker_row).Value = Percent_change
              Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
'Reset the row counter. Add +1 to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

'Reset volume of stock to zero
              Stockvolume = 0

'Reset opening price
              Open_price = Cells(i + 1, 3)
            
            Else
              
'Add the volume of stock
             Stockvolume = tickervolume + Cells(i, 7).Value

            
            End If
        
        Next i

'Conditional formatting

    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
'Color code yearly change
    
    For i = 2 To lastrow_summary_table
    
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i

End Sub

