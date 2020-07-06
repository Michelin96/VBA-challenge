Attribute VB_Name = "Module1"
Sub stock_analysis()

Dim ticker As String
Dim change As Double
Dim pct_change As Double
Dim total_stock As Single
Dim summary_row As Integer
Dim new_row As Long
Dim last_row As Long
Dim ws_count As Integer
Dim new_ws As Integer

'outer for loop iterates through the worksheets
ws_count = ActiveWorkbook.Worksheets.Count
For new_ws = 1 To ws_count

'this creates a new summary table on each worksheet
Worksheets(new_ws).Cells(1, 11).Value = "Ticker"
Worksheets(new_ws).Cells(1, 12).Value = "Yearly Change"
Worksheets(new_ws).Cells(1, 13).Value = "Percent Change"
Worksheets(new_ws).Cells(1, 14).Value = "Total Stock Volume"
'expand the columns to fit the labels and data
Worksheets(new_ws).Columns("A:N").AutoFit
summary_row = 2
total_stock = 0

'get the value for open in the second row of the new worksheet
open_pts = Worksheets(new_ws).Cells(2, 3).Value2

'find last row in worksheet
last_row = Worksheets(new_ws).Cells(Rows.Count, 1).End(xlUp).Row

    'inner for loop interates through the active worksheet rows
    For new_row = 2 To last_row
    
    'check to see if the ticker symbol changed and if it has...
    If Worksheets(new_ws).Cells(new_row + 1, 1).Value <> Worksheets(new_ws).Cells(new_row, 1).Value Then
  
        'add the last row of that stock's volume to the total stock
        total_stock = total_stock + Worksheets(new_ws).Cells(new_row, 7).Value
        
        'get the current ticker symbol and put it in the summary table
        ticker = Worksheets(new_ws).Cells(new_row, 1).Value
        Worksheets(new_ws).Range("K" & summary_row).Value = ticker
        
        'put the total ticker stock volume in the summary table
        Worksheets(new_ws).Range("N" & summary_row).Value = total_stock
        'reset total volume to 0 because the symbol is changing
        total_stock = 0
      
       'get the close of the last row for the current ticker and put it in close_pts
        close_pts = Worksheets(new_ws).Cells(new_row, 6).Value

        'calculate the yearly change and put it in the summary table
        change = close_pts - open_pts
        Worksheets(new_ws).Range("L" & summary_row).Value = change
            If change < 0 Then
                'format neg change vaules to red
                Worksheets(new_ws).Range("L" & summary_row).Interior.ColorIndex = 3
            Else
                'format pos change values to green
                Worksheets(new_ws).Range("L" & summary_row).Interior.ColorIndex = 4
            End If
        
        'calculate percent change and put it in the sumary table
            If (change <> 0 And open_pts <> 0) Then
                pct_change = Round((change / open_pts) * 100, 2)
                Worksheets(new_ws).Range("M" & summary_row).Value = pct_change & "%"
            Else
                'the vaules in change or open_pts are 0 and a pct_change cannot be calculated
                Worksheets(new_ws).Range("M" & summary_row).Value = "0"
            End If
            
        'get the open value of the first row in the next ticker symbol
        open_pts = Worksheets(new_ws).Cells(new_row + 1, 3).Value
     
        'Add one to the summary table row
        summary_row = summary_row + 1
        
    Else
    'this is the only thing that is done until the ticker changes
    'sum the stock volume
    total_stock = total_stock + Worksheets(new_ws).Cells(new_row, 7).Value
    
    End If
    
    Next new_row

Next new_ws

End Sub
