'This exercise uses VBA to evaluate large sets of stock data with the ability to go through each sheet of a workbook.
'Calculations are performed to evaluate which stock had the greatest and smallest stock price increase in that year.
'VBA was used to keep track of the total stock volume of each stock for the year.

Sub Master_Loop()
'Repeats into each sheet within the workbook
    For Each ws In Worksheets
        ws.Activate
        'Running  Sub programs/procedures VBA homework and VBA challenge to execute
            Call VBA_homework
            Call VBA_challenge
     Next ws
End Sub


Sub VBA_homework()
  
    'Declares variables
        Dim ticker As String
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_stock As Double
   'Keeps record of the row used in the summary table
        Dim summary_row As Integer

    'Inserts header row in the summary table
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Perecent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("I:L").Columns.AutoFit

    'Determines last row
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Sets the first row of the summary table row to 2 to account for headers
        summary_row = 2
   
    'Loops through stock data to find/calculate values for variables defined
        For Row = 2 To lastrow
        
            'Identifies opening price and ticker by comparing current and previous rows
                If Cells(Row, 1) <> Cells(Row - 1, 1) Then
                    open_price = Cells(Row, 3).Value
                    ticker = Cells(Row, 1).Value
                    Cells(summary_row, 9) = ticker
                End If
        
            'Identifies closing price and calculates for pricing changes by comparing current and next rows
                If Cells(Row + 1, 1) <> Cells(Row, 1) Then
                    close_price = Cells(Row, 6).Value
        
                    'Calculates for yearly change of each stock
                    yearly_change = close_price - open_price
                    Cells(summary_row, 10) = yearly_change
                    
                
                     'Changes cells with stock price increases to green and price drops to red
                        If yearly_change > 0 Then
                            Cells(summary_row, 10).Interior.ColorIndex = 4
                        Else
                            Cells(summary_row, 10).Interior.ColorIndex = 3
                        End If
                        
                    'Calculates the increase/decrease in stock price in percentages
                    'If opening price is 0, the percentage change in stock price will be calculated dependent on the closing price.
                        If open_price = 0 And close_price <> 0 Then
                            percent_change = 1
                        ElseIf (open_price = 0 And close_price = 0) Then
                            percent_change = 0
                        Else
                            percent_change = (yearly_change / open_price)
                        End If
          
                    'Converts decimals to display as percentages
                    Cells(summary_row, 11) = Format(percent_change, "percent")
              
                    'Adds the total amount of stock and displays in summary table
                    total_stock = total_stock + Cells(Row, 7).Value
                    Cells(summary_row, 12) = total_stock
                    
                    'Displays on next row for summary table
                    summary_row = summary_row + 1
                    
                    'Reset stocks to 0 when identifying new ticker/stock
                    total_stock = 0
            
                Else
                
                    'Continues to add stock volumes within same ticker/stock
                    total_stock = total_stock + Cells(Row, 7).Value
            
                End If
    
        Next Row

End Sub


Sub VBA_challenge()

'Declares variables
Dim max_inc As Double
Dim min_inc As Double
Dim total As Double
Dim ticker_challenge As String
    
'Places headers for challenge table
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    
'Counts total number of rows in summary table
    lastrow = Cells(Rows.Count, 11).End(xlUp).Row

'Loops to find greatest and smallest % increase and writes into the challenge table with corresponding ticker
    For Row = 2 To lastrow
    
        'Identifies the max and min percent changes along with larges stock volume and displays in column Q
            max_inc = Application.WorksheetFunction.Max(Range("k:k"))
            min_inc = Application.WorksheetFunction.Min(Range("k:k"))
            total = Application.WorksheetFunction.Max(Range("l:l"))
            
        'Displays the greatest % increase and decrease as percentages
            Range("Q2") = Format(max_inc, "percent")
            Range("Q2").Value = max_inc
            Range("Q3") = Format(min_inc, "percent")
            Range("Q3").Value = min_inc
            Range("Q4").Value = total
        
            'Finds corresponding ticker with each value of column Q
                If Cells(Row, 11) = max_inc Then
                    ticker_challenge = Cells(Row, 9).Value
                    Range("P2") = ticker_challenge
                End If
                
                If Cells(Row, 11) = min_inc Then
                    ticker_challenge = Cells(Row, 9).Value
                    Range("p3") = ticker_challenge
                End If
            
                If Cells(Row, 12) = total Then
                    ticker_challenge = Cells(Row, 9).Value
                    Range("p4") = ticker_challenge
                End If
        
    Next Row
    
Range("O:Q").Columns.AutoFit

End Sub




