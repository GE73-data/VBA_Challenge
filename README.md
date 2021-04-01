# VBA_Challenge
Wall Street Stock Analysis Homework
---
Sub Gloria_Stock_Market_Analysis():
    'Declare Variable
    Dim Ticker As String
    Dim current_row As Long
    Dim last_row As Long
    Dim summary_row As Integer
    Dim total_volume As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_price As Double
    Dim percent_change As Double
    Dim first_sale_price As Long
    Dim start_row As Long
    
    
    'Initialize Variables
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    summary_row = 2
    total_volume = 0
    open_price = 0
    close_price = 0
    yearly_change = 0
    percent_change = 0
    first_sale_price = 2
    start_row = 2
   

    
    'Create Headers in Summary Table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Columns("I:L").EntireColumn.AutoFit
    Columns("I:L").Font.Bold = True
    
    'Iterate through the worksheet from row 2 to last_row
    For current_row = 2 To last_row
    
        If Cells(current_row + 1, 1) <> Cells(current_row, 1) Then
                  
        'set ticker
        Ticker = Cells(current_row, 1).Value
        
        total_volume = total_volume + Cells(current_row, 7).Value
        
        'Print ticker to summary table
            Range("I" & summary_row).Value = Cells(current_row, 1).Value
            
         If total_volume = 0 Then
         Range("J" & summary_row).Value = 0
         Range("K" & summary_row).Value = 0
         Range("L" & summary_row).Value = 0
         
         Range("J" & summary_row).Interior.ColorIndex = 8
         
         Else
         If Cells(start_row, 3).Value = 0 Then
         
         For first_sale_price = start_row To current_row
         
         If Cells(first_sale_price, 3).Value <> 0 Then
         start_row = first_sale_price
         
         Exit For
         
         End If
         
         Next first_sale_price
         
                 
         End If
         
         'Print total stock volume
            Range("L" & summary_row).Value = total_volume
            
        'Yearly change from open price to closing
        open_price = Cells(start_row, 3).Value
        close_price = Cells(current_row, 6).Value
        yearly_change = close_price - open_price
        
        'Print yearly change
        Range("J" & summary_row).Value = yearly_change
        
            
        'percent change calculation
        percent_change = (yearly_change / open_price)
        
        'Print percent_change
        Range("K" & summary_row).Value = percent_change
        
        Range("K" & summary_row).NumberFormat = "0.00%"
        
    'Format highlighting positive change green and negative change red
        
        If Range("J" & summary_row).Value > 0 Then
        Range("J" & summary_row).Interior.ColorIndex = 4
        
        Else
        Range("J" & summary_row).Interior.ColorIndex = 3
        
        End If
        
        End If
                
            'Reset ticker variables for next ticker
            total_volume = 0
            summary_row = summary_row + 1
        
        
        'Reset start row
        start_row = current_row + 1
        
        Else
        total_volume = total_volume + Cells(current_row, 7).Value
        
        
        End If
            
    Next current_row
End Sub


