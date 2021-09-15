Attribute VB_Name = "Module1"
Sub Run_analysis()

'Dimension initial variables present in spreadsheet
Dim ticker_symbol As String
Dim Open_Value As Double
Dim Close_Value As Double
Dim open_price_column As Integer
Dim volume_column As Integer


'dimension calculated variables in spreadsheet
Dim Yearly_Change As Double
Dim Percent_Change As Double

Dim new_ticker_column As Integer
Dim new_yearly_change_column As Integer
Dim new_percent_change_column As Integer
Dim new_total_stock_volume_column As Integer

 
'created variables to keep track of cloumn numbers;readability
ticker_column = 1
date_column = 2
open_price_column = 3
close_price_column = 6
volume_column = 7

  
  'keep track of new column numbers
new_ticker_column = 9
new_yearly_change_column = 10
new_percent_change_column = 11
new_total_stock_volume_column = 12

Dim current_volume As Double
Dim print_row As Double
Dim running_volume_total As Double

'initialize values
print_row = 2
running_volume_total = 0
Yearly_Change = 0
Open_Value = 0
Close_Value = 0
Dim ws As Worksheet



    For Each ws In Worksheets
    
    
    
        'add headers for newly calculated values
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
        'find last row in each worksheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
      
        
        Open_Value = ws.Cells(2, 3).Value
        
        For i = 2 To lastrow
        
        
        
        ' Loop through rows to see if the ticker symbol is the same, if not...
            If ws.Cells(i + 1, ticker_column).Value <> ws.Cells(i, ticker_column).Value Then
    
                
                
                    'finds opening price of the year for that ticker
                    'If ws.Cells(i, date_column).Value = 20161230 Then
                    
                    Close_Value = ws.Cells(i, close_price_column).Value
                    'End If
                    
                
                'calculate yearly change
                Yearly_Change = Close_Value - Open_Value
                   
                'calculate percent change from opening price to close prince of that year
                
                If Open_Value <> 0 Then
                
                    Percent_Change = (Yearly_Change / Open_Value) * 100
                    
                Else
                
                    Percent_Change = 0
                
                End If
                    
                running_volume_total = running_volume_total + ws.Cells(i, volume_column).Value
                
                'print the current ticker symbol to the new column labeled Ticker
                ws.Cells(print_row, new_ticker_column).Value = ticker_symbol
                 
                'print the runnning Total to the new column labeled Total Stock Volume
                ws.Cells(print_row, 12).Value = running_volume_total
                
                 'print the percent change to the Percent Change column
                ws.Cells(print_row, new_percent_change_column).Value = Percent_Change
                
                
        
                'print the yearly change value to column labeled Total Stock Volume
                ws.Cells(print_row, new_yearly_change_column).Value = Yearly_Change
                
                    If Yearly_Change < 0 Then
                    
                        ws.Cells(print_row, new_yearly_change_column).Interior.ColorIndex = 3
                        
                    Else
                        ws.Cells(print_row, new_yearly_change_column).Interior.ColorIndex = 4
                        
                    End If
                    
                Open_Value = ws.Cells(i + 1, open_price_column)
                
                
                'reset all totals
                running_volume_total = 0
                Yearly_Change = 0
                Percent_Change = 0
                'opening_year_price = 0
                'closing_year_price = 0
                'increment print row for calculated values
                print_row = print_row + 1
            
            
            'if the ticker symbol remains unchanged
            Else
                
                'assigns the current ticker to variable ticker_symbol
                ticker_symbol = ws.Cells(i, ticker_column)
                
                'add that day's volume to running total
                running_volume_total = running_volume_total + ws.Cells(i, 7).Value
                
                    
                
            End If
        
        Next i
        
    'reset print row counter
    print_row = 2

    Next ws
            
End Sub
               
               

        

  'Next i
  


