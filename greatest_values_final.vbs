Sub greatest_values()

Dim ticker_vol As String
Dim ticker_inc As String
Dim ticker_dec As String
Dim summary_row As Integer
Dim last_row As Long
Dim ws_count As Integer
Dim new_ws As Integer
Dim greatest_volume As Single
Dim greatest_increase As Double
Dim greatest_decrease As Double

'Cycle through the worksheets
ws_count = ActiveWorkbook.Worksheets.Count
For new_ws = 1 To ws_count

    'create a new summary table for the greatest values on each worksheet and
    'expand the columns to fit the labels and data
    Worksheets(new_ws).Columns("Q:S").AutoFit
    'rows
    Worksheets(new_ws).Cells(2, 17).Value = "Greatest % Increase"
    Worksheets(new_ws).Cells(3, 17).Value = "Greatest % Decrease"
    Worksheets(new_ws).Cells(4, 17).Value = "Greatest Volume"
    'columns
    Worksheets(new_ws).Cells(1, 18).Value = "Ticker"
    Worksheets(new_ws).Cells(1, 19).Value = "Value"
    
    'get the last row on the summary table
    last_row = Worksheets(new_ws).Range("K" & Rows.Count).End(xlUp).Row
    
    'resets the vaules on the new worksheet
    summary_row = 2
    greatest_volume = 0
    greatest_increase = 0
    greatest_decrease = 0
        
        'cycle through the rows on the summary tables
        For summary_row = 2 To last_row
        
            'compare each row to the current greatest increase and store the larger value and its symbol
            If Worksheets(new_ws).Cells(summary_row, 13).Value > greatest_increase Then
                greatest_increase = Worksheets(new_ws).Cells(summary_row, 13).Value
                ticker_inc = Worksheets(new_ws).Cells(summary_row, 11).Value
            End If
            
            'compare each row to the current greatest decrease and store the larger value and its symbol
            If Worksheets(new_ws).Cells(summary_row, 13).Value < greatest_decrease Then
                greatest_decrease = Worksheets(new_ws).Cells(summary_row, 13).Value
                ticker_dec = Worksheets(new_ws).Cells(summary_row, 11).Value
            End If
    
            'compare each row to the current greatest volume and store the larger value and its symbol
            If Worksheets(new_ws).Cells(summary_row, 14).Value > greatest_volume Then
                greatest_volume = Worksheets(new_ws).Cells(summary_row, 14).Value
                ticker_vol = Worksheets(new_ws).Cells(summary_row, 11).Value
            End If
            
        Next summary_row
        
        'after all the summary rows have been evaluated, record the values in the greatest summary table
        Worksheets(new_ws).Range("R2").Value = ticker_inc
        Worksheets(new_ws).Range("R3").Value = ticker_dec
        Worksheets(new_ws).Range("R4").Value = ticker_vol
        Worksheets(new_ws).Range("S2").Value = Round((greatest_increase) * 100, 2) & "%"
        Worksheets(new_ws).Range("S3").Value = Round((greatest_decrease) * 100, 2) & "%"
        Worksheets(new_ws).Range("S4").Value = greatest_volume
        'expand the columns to fit the label and data
        Worksheets(new_ws).Columns("Q:S").AutoFit
        
   Next new_ws

End Sub

