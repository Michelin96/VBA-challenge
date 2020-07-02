Attribute VB_Name = "Module1"
Sub stock_analysis()

Dim ticker As String
Dim change As Double
Dim pct_change As Double
Dim total_stock As Single
Dim summary_row As Integer
Dim next_row As Long
Dim last_row As Long
Dim ws_count As Integer
Dim next_ws As Integer

'outer for loop iterates through the worksheets
ws_count = ActiveWorkbook.Worksheets.Count
For next_ws = 1 To ws_count
'Debug.Print (ActiveWorkbook.Worksheets(next_ws).Name)

'this creates a new summary table on each worksheet
Worksheets(next_ws).Cells(1, 11).Value = "Ticker"
Cells(1, 12).Value = "Yearly Change"
Cells(1, 13).Value = "Percent Change"
Cells(1, 14).Value = "Total Stock Volume"
summary_row = 2
total_stock = 0

'get the value for open in the second row of the new worksheet
open_pts = Cells(2, 3).Value2

'find last row in worksheet
last_row = Cells(Rows.Count, 1).End(xlUp).Row

    'inner for loop interates through the active worksheet rows
    For next_row = 2 To last_row
    
    'check to see if the ticker symbol changed
    If Cells(next_row + 1, 1).Value <> Cells(next_row, 1).Value Then
  
        'add the last row of that stock volume to the total stock
        total_stock = total_stock + Cells(next_row, 7).Value
        
        'get the current ticker symbol and put it in the summary table
        ticker = Cells(next_row, 1).Value
        Range("K" & summary_row).Value = ticker
        
        'put the total ticker stock in the summary table
        Range("N" & summary_row).Value = total_stock
        'reset total volume to 0 when the ticker symbol changes
        total_stock = 0
      
       'get the close of the last row and put it in close_pts
        close_pts = Cells(next_row, 6).Value

        'calculate the yearly change and put it in the summary table
        change = close_pts - open_pts
        Range("L" & summary_row).Value = change
            If change < 0 Then
                'format neg change vaules to red
                Range("L" & summary_row).Interior.ColorIndex = 3
            Else
                'format pos change values to green
                Range("L" & summary_row).Interior.ColorIndex = 4
            End If
        
        'calculate percent change and put it in the sumary table
        pct_change = Round((change / open_pts) * 100, 2)
        Range("M" & summary_row).Value = pct_change & "%"
        
        'get the open value of the first row in the next ticker symbol
        open_pts = Cells(next_row + 1, 3).Value
     
        'Add one to the summary table row
        summary_row = summary_row + 1
        
    Else

   'sum the stock volume
    total_stock = total_stock + Cells(next_row, 7).Value
    
    End If
    
    Next next_row

Next next_ws
End Sub
Sub stock_analysis()

Dim ticker As String
Dim change As Double
Dim pct_change As Double
Dim total_stock As Single
Dim summary_row As Integer
Dim next_row As Long
Dim last_row As Long
Dim ws_count As Integer
Dim next_ws As Integer

'outer for loop iterates through the worksheets
ws_count = ActiveWorkbook.Worksheets.Count
For next_ws = 1 To ws_count

'this creates a new summary table on each worksheet
Worksheets(next_ws).Cells(1, 11).Value = "Ticker"
Cells(1, 12).Value = "Yearly Change"
Cells(1, 13).Value = "Percent Change"
Cells(1, 14).Value = "Total Stock Volume"
summary_row = 2
total_stock = 0

'get the value for open in the second row of the new worksheet
open_pts = Cells(2, 3).Value2

'find last row in worksheet
last_row = Cells(Rows.Count, 1).End(xlUp).Row

    'inner for loop interates through the active worksheet rows
    For next_row = 2 To last_row
    
    'check to see if the ticker symbol changed
    If Cells(next_row + 1, 1).Value <> Cells(next_row, 1).Value Then
  
        'add the last row of that stock volume to the total stock
        total_stock = total_stock + Cells(next_row, 7).Value
        
        'get the current ticker symbol and put it in the summary table
        ticker = Cells(next_row, 1).Value
        Range("K" & summary_row).Value = ticker
        
        'put the total ticker stock in the summary table
        Range("N" & summary_row).Value = total_stock
        'reset total volume to 0 when the ticker symbol changes
        total_stock = 0
      
       'get the close of the last row and put it in close_pts
        close_pts = Cells(next_row, 6).Value

        'calculate the yearly change and put it in the summary table
        change = close_pts - open_pts
        Range("L" & summary_row).Value = change
            If change < 0 Then
                'format neg change vaules to red
                Range("L" & summary_row).Interior.ColorIndex = 3
            Else
                'format pos change values to green
                Range("L" & summary_row).Interior.ColorIndex = 4
            End If
        
        'calculate percent change and put it in the sumary table
        pct_change = Round((change / open_pts) * 100, 2)
        Range("M" & summary_row).Value = pct_change & "%"
        
        'get the open value of the first row in the next ticker symbol
        open_pts = Cells(next_row + 1, 3).Value
     
        'Add one to the summary table row
        summary_row = summary_row + 1
        
    Else

   'sum the stock volume
    total_stock = total_stock + Cells(next_row, 7).Value
    
    End If
    
    Next next_row

Next next_ws
End Sub

