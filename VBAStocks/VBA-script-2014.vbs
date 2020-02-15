'1. list all stock Ticker symbols
'2. for each stock, yearly change from opening price at beginning of given year to closing price at end of that year
'3. for each stock, percent change " "
'4. for each stock, total stock volume
'5. conditional formatting of yearly change


Sub Ticker3()

    'set initial variable for holding ticker symbol
    Dim Ticker As String
    
    'set initial variable for open price
    Dim OpenPrice As Double
    
    'set initial variable for close price
    Dim ClosePrice As Double
    
    'set initial variable for holding yearly change
    Dim Yearly_Change As Double
            
    'set initial variable for holding percent change
    Dim Percent_Change As Double
    
    'set variable for holding percent change with percent formatting
    Dim Percent_Change_Formatted As String
        
    'set initial variable for holding total stock volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
            
    'keep track of location for ea ticker symbol in summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'name column I "Ticker"
    Range("I1").Value = "Ticker"
    
    'name column J "Yearly Change"
    Range("J1").Value = "Yearly Change"
    
    'name column K "Percent Change"
    Range("K1").Value = "Percent Change"
    
    'name column L "Total Stock Volume"
    Range("L1").Value = "Total Stock Volume"
    
    'declare ws as a worksheet object variable
    'Dim ws As Worksheet
    
    'loop through all sheets
    'For Each ws In Worksheets
    
        'find the last row of each worksheet
        'lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
        'loop through all rows of data
        For i = 2 To lastrow
        
                                   
            'if we're in the first row of a ticker symbol where it's different from row above...
            If (Cells(i - 1, 1).Value <> Cells(i, 1).Value) Then
            
                'set open price
                'OpenPrice = ws.Cells(i, 3).Value
                OpenPrice = Cells(i, 3).Value
                  
                
            'check if we are still within the same ticker symbol. if not..
            ElseIf (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
            
                'set the ticker name
                'Ticker = ws.Cells(i, 1).Value
                Ticker = Cells(i, 1).Value
                
                'print the ticker name in the summary table
                Range("I" & Summary_Table_Row).Value = Ticker
                           
                'set close price
                'ClosePrice = ws.Cells(i, 6).Value
                ClosePrice = Cells(i, 6).Value
                
                'add to the yearly change
                Yearly_Change = ClosePrice - OpenPrice
                                
                'print the yearly change to the summary table
                Range("J" & Summary_Table_Row).Value = Yearly_Change
        
                'add to percent change if openprice does not equal zero
                If OpenPrice <> 0 Then
                
                    Percent_Change = Yearly_Change / OpenPrice
                
                Else
                
                End If
                
                'format percent change as percent
                Percent_Change_Formatted = FormatPercent(Percent_Change, 2)
                
                'print the percent change to the summary table
                Range("K" & Summary_Table_Row).Value = Percent_Change_Formatted

                'add to the total stock volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
                'print the total stock volume to the summary table
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                                
                'add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                               
                'reset the total stock volume
                Total_Stock_Volume = 0
                
            'if the cell immediately following a row is the same ticker symbol..
            Else
            
                'add to the total stock volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

            End If
            
            'conditional formatting of yearly change
            'if yearly change is positive
            If (Cells(i, 10).Value > 0) Then
                    
                'highlight green
                Cells(i, 10).Interior.ColorIndex = 4
                    
            'if yearly change is negative
            ElseIf (Cells(i, 10).Value < 0) Then
                    
                'highlight red
                Cells(i, 10).Interior.ColorIndex = 3
                    
            End If
                            
             
        Next i
        
       
    'Next ws
         
End Sub
