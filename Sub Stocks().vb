Sub Stocks()

    ' Set an initial variable for ticker
    Dim ticker As String
  
    ' Set an initial variable for open price
    Dim open_price As Double
  
    'Set an initial variable for close price
    Dim close_price As Double
    
    'Set an initial variable for yearly change
    Dim yearly_change As Double
   
    'Set an initial variable for percent change
    Dim pct_change As Double
    
    ' Set an initial variable for total volume
    Dim volume As Double
    volume = 0
    
    ' Set an initial variable for ticker with greatest % increase, % decrease, and volume
    Dim greatest_volume_ticker As String
    Dim greatest_pct_ticker As String
    Dim greatest_pct_decrease_ticker As String
    
    ' Set an initial variable for ticker with greatest % increase, % decrease and volume
    Dim greatest_volue As Double
    Dim greatest_increase_pct As Double
    Dim greatest_decrease_pct As Double
    
    ' Set varible for worksheets
    Dim ws As Worksheet
  
    ' Keep track of the location for each ticker in a summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Additional Functionality Summary Table
    Dim Summary_Greatest As Integer
    Summary_Greatest = 2
    
    ' Set variable for loop in all worksheets
    For Each ws In Worksheets
   
   ' Label column headers for Summary Table
    ws.Cells(1, 11).Value = "Ticker"
    ws.Cells(1, 12).Value = "Yearly Change"
    ws.Cells(1, 13).Value = "Percent Change"
    ws.Cells(1, 14).Value = "Total Volume"
    ws.Cells(1, 16).Value = "Additional Functionality"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    
    ' Find last row in column A
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    open_price = ws.Cells(2, 3).Value
    
    'Set values to original starting points. 99999999 is used as default high number for comparision.
    greatest_increase_pct = 0
    greatest_decrease_pct = 99999999
    greatest_volume = 0
    
    ' Start Loop
        For i = 2 To lastrow
        
            ' Aggreate volume totals
                volume = volume + ws.Cells(i, 7).Value
                
            ' Find close price for each ticker
                close_price = ws.Cells(i, 6).Value
                
            ' Calculate the Yearly Change
                yearly_change = close_price - open_price
                
                ' Calculate the Percent Change
                If open_price > 0 Then
                    pct_change = (close_price - open_price) / open_price
                Else
                    pct_change = 0
                End If
              
            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                ' Set the ticker name
                ticker = ws.Cells(i, 1).Value
                  
                ' Find open price for each ticker
                open_price = ws.Cells(i + 1, 3).Value
                
                If pct_change > greatest_increase_pct Then
                    greatest_increase_pct = pct_change
                    greatest_pct_ticker = ws.Cells(i, 1).Value
                End If
                
                If pct_change < greatest_decrease_pct Then
                    greatest_decrease_pct = pct_change
                    greatest_pct_decrease_ticker = ws.Cells(i, 1).Value
                End If
                
                If volume > greatest_volume Then
                    greatest_volume = volume
                    greatest_volume_ticker = ws.Cells(i, 1).Value
                End If
                
                
            ' Print the Summary Table
            ' Omitting the ws.Cells will put all records on first tab
            ws.Cells(Summary_Table_Row, 11).Value = ticker
            ws.Cells(Summary_Table_Row, 12).Value = yearly_change
            ws.Cells(Summary_Table_Row, 13).Value = pct_change
            ws.Cells(Summary_Table_Row, 14).Value = volume
                    
      
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            Summary_Greatest = Summary_Greatest + 1
            
            ' Reset the volume Total
            volume = 0
            yearly_change = 0
            
            End If
            
        Next i
        
      Summary_Table_Row = 2
      Summary_Greatest = 2
      
      
      ws.Cells(4, 17).Value = greatest_volume_ticker
      ws.Cells(4, 18).Value = greatest_volume
        ws.Cells(4, 18).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
      ws.Cells(2, 17).Value = greatest_pct_ticker
      ws.Cells(2, 18).Value = greatest_increase_pct
        ws.Cells(2, 18).NumberFormat = "0.00%"
      ws.Cells(3, 17).Value = greatest_pct_decrease_ticker
      ws.Cells(3, 18).Value = greatest_decrease_pct
        ws.Cells(3, 18).NumberFormat = "0.00%"
       
 'Conditional Formatting
 '-------------------------------------------------------------------
 
    year_format = ws.Cells(Rows.Count, 12).End(xlUp).Row
    For y = 2 To year_format
 
 ' For Yearly Change colume, set values to color Red(3) if less than 0, else set color to Green(4)
        If ws.Cells(y, 12).Value < 0 Then
            
        ws.Cells(y, 12).Interior.ColorIndex = 3
            
        Else
        ws.Cells(y, 12).Interior.ColorIndex = 4
        End If

    Next y
    
    percent_format = ws.Cells(Rows.Count, 13).End(xlUp).Row
    For p = 2 To percent_format
 
 ' For Percent Change column, set values to color Red(3) if less than 0, else set color to Green(4)
        If ws.Cells(p, 13).Value < 0 Then
            
        ws.Cells(p, 13).Interior.ColorIndex = 3
            
        Else
        ws.Cells(p, 13).Interior.ColorIndex = 4
        End If

    Next p
  
  
Next ws

End Sub