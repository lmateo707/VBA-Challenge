# VBA-Challenge
VBA_Homework_Mateo

'Instructions
    'create a script that will loop through all the stocks for one year and output the following information
    'the ticker symbol
    'yearly change from opening price at the beginning of the given year to the closing price at the end of the year
    'the percent change from opening price at the beginning of a given year to the closing price of that year
    'the total stock volume
'Instructions
    'create a script that will loop through all the stocks for one year and output the following information
    'the ticker symbol
    'yearly change from opening price at the beginning of the given year to the closing price at the end of the year
    'the percent change from opening price at the beginning of a given year to the closing price of that year
    'the total stock volume

'conditional formatting
    'highlight positive change in the green and negative change in the red


Sub Multi_year_stock_data()

    'Loop through all the sheets
    For Each ws In Worksheets
    
     Dim Worksheet As String
     
    
    'Create variable to hold Ticker Symbol, Yearly Change, Percent Change, Total Stock Volume
        Dim Ticker_Symbol As String
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_stock_volume As Double
        
        
    'Location for ticker symbol in the table
        Dim table_row As Integer
        table_row = 2
    
    'Loop through all stocks
        For i = 2 To 70926
    
    
       
        'identify when there is a next ticker symbol
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            
        
        
        'set the ticker symbol name
            Ticker_Symbol = Cells(i, 1).Value
         
            yearly_change = Cells(i, 6).Value - Cells(i, 3).Value
        
            percent_change = Cells(i, 4).Value - Cells(i, 5).Value
        
            total_stock_volume = yearly_change * percent_change
        
        
        'print ticker symbol in table row
            Range("I" & table_row).Value = Ticker_Symbol
            Range("J" & table_row).Value = yearly_change
            Range("K" & table_row).Value = percent_change
            Range("K" & table_row).Style = "percent"
            
            Range("L" & table_row).Value = total_stock_volume
        
            table_row = table_row + 1
        
            yearly_change = 0
        
        
        
        
            Else
        
        
            yearly_change = yearly_change + Cells(i, 3).Value
    
            End If
     
     Next i
     
     'yearly change
     For i = 2 To 262
     
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        yearly_change = Cells(i, 6).Value - Cells(i, 3).Value
        Range("J" & table_row).Value = yearly_change
        
     
     End If
     
     
     Next i

    
        
    
    
    Next i
    
     
     
     
     Next ws
     
     
     
    
    End Sub
    
   Sub color_change()

    'color change for yearly change
     
    If Cells(10, 290).Value > 0 Then
    Range("J2:J290").Interior.ColorIndex = 4
    
    ElseIf Cells(10, 290).Value <= 0 Then
    
    Range("J2:J290").Interior.ColorIndex = 3
    
    
    
    
    End If
    
    
    
     
   End Sub
   
    
    
    
  
    

     
     
     
