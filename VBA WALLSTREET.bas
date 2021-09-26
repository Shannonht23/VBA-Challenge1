Attribute VB_Name = "Module1"
Sub wallstreet():


    'set an initial variable for holding the ticker symbol
    Dim ticker_symbol As String
    
    'set an initial variable for holding the total_stock_volume
    Dim total_stock_volume As LongLong
    total_stock_volume = 0
    
    'set an initial variable for holding the percent change
    Dim percent_change As Double
    Dim percent_change_cell As String
    
    'set an inital variable for holding the yearly change
    Dim yearly_change As Double
    
    'set an inital variable for holding opening price
    
    Dim opening_price As Double
    
    'Set initial variable for holding closing price
    Dim closing_price As Double
    
    'set initial variable for minimum date
    Dim minimum_date As LongLong
    
    'set initail variable for maximum date
    Dim maximum_date As LongLong
    
    Dim max_row_count As LongLong
    
    Dim Summary_Table_Row As Integer
    
    ' Some for loop , that loops through each sheet
    
    Dim color_index As Integer
    
    
    For Each ws In Worksheets
    
    
        'keep track of the location for each ticker symbol in the summary table
         Summary_Table_Row = 2
           
         
         'counts the number of rows
         
         max_row_count = ws.Cells(Rows.Count, 1).End(xlUp).Row
         minimum_date = 99999999
         maxiumum_date = 0
         yearly_change = 0
         opening_price = 0
         closing_price = 0
         total_stock_volume = 0
         percent_change = 0
         
         'Label Summary Columns for sheet summary table
         
         ws.Range("j1").Value = "Ticker Symbol"
         ws.Range("k1").Value = "Yearly change"
         ws.Range("l1").Value = "Percentage Change"
         ws.Range("m1").Value = "Total Stock Volume"
         
         
         
             'Loop through all ticker symbols
             For i = 2 To max_row_count
             
                 'check if we are still within the same ticker symbol, if its not.
                 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             
                     If ws.Cells(i, 2).Value < minimum_date Then
                         minimum_date = ws.Cells(i, 2).Value
                         opening_price = ws.Cells(i, 3)
                     End If
                           
                     If ws.Cells(i, 2).Value > maximum_date Then
                         maximum_date = ws.Cells(i, 2).Value
                         closing_price = ws.Cells(i, 6)
                     End If
                     
                     'set the ticker symbol
                     ticker_symbol = ws.Cells(i, 1).Value
                 
                     'add to the total_stock_volume
                     total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                 
                     'Add yearly change
                     
                     yearly_change = closing_price - opening_price
                     
                     
                     If opening_price = 0 Then
                     opening_price = 1
                     End If
                     
                     percent_change = yearly_change / opening_price
                     percent_change_cell = FormatPercent(percent_change, 2)
                   
                     'Print the ticker symbol in the summary table
                     ws.Range("j" & Summary_Table_Row).Value = ticker_symbol
                 
                           
                     'Print the yearly_change in the summary table
                     ws.Range("K" & Summary_Table_Row).Value = yearly_change
                     
                     If yearly_change >= 0 Then
                        color_index = 50
                     Else
                        color_index = 53
                     End If
                     
                     ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = color_index
                     
                     
                     
                     'Print the percent_change in the summary table
                     ws.Range("l" & Summary_Table_Row).Value = percent_change_cell
                     
                     
                     'Print the percent_change in the summary table
                     ws.Range("M" & Summary_Table_Row).Value = total_stock_volume
                     
                     
                     'add one to the summary table row
                      Summary_Table_Row = Summary_Table_Row + 1
                      
                     
                     'reset the stock volume total
                     total_stock_volume = 0
                     minimum_date = 99999999
                     maximum_date = 0
                     opening_price = 0
                     closing_price = 0
                     percent_change = 0
                     yearly_change = 0
             
                 Else
             
                     total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                     If ws.Cells(i, 2).Value < minimum_date Then
                         minimum_date = ws.Cells(i, 2).Value
                         opening_price = ws.Cells(i, 3)
                     End If
                     
                     If ws.Cells(i, 2).Value > maximum_date Then
                         maximum_date = ws.Cells(i, 2).Value
                         closing_price = ws.Cells(i, 6)
                     End If
                     
                     
                         
                     
                     
             
             
                 End If
             Next i
    Next ws




End Sub
