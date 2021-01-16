Attribute VB_Name = "Module1"
Sub StockCalculations()

    
    Dim current_ticker As String
    Dim next_ticker As String
    
    'r is the row counter and c is the column counter
    Dim r As Double
    Dim c As Integer
    
    'Dim ws As Worksheet
    Dim last_row As Double
    
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim percentage_change_value As Double
    
    'declare greatest variables
    Dim greatest_percent_increase_ticker As String
    Dim greatest_percent_decrease_ticker As String
    Dim greatest_total_volume_ticker As String
    
    Dim greatest_percent_increase_value As Double
    Dim greatest_percent_decrease_value As Double
    Dim greatest_total_volume_value As Double
    
    
    Dim i As Double
        
    Dim total_volume As Double
    Dim current_volume As Double
    'Set current_sheet = Worksheets("A")
    
    
    Dim last_sheet As Double
    
    Dim ws As Worksheet
    'loop through each sheet in the workbook
    
    last_sheet = Sheets.Count
    
    For Each ws In Worksheets
    
        'Debug.Print ws.Name
        'at the start of each worksheet initialize the output row counter to the start
        i = 2
        
        'insert output column headers on every worksheet
        ws.Cells(1, 10).Value = "Ticker"
        ws.Columns("J").ColumnWidth = 8
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Columns("K").ColumnWidth = 12
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Columns("L").ColumnWidth = 13
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Columns("M").ColumnWidth = 16
        
        'insert Greatest data column and row labels
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Columns("R").ColumnWidth = 12
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Columns("P").ColumnWidth = 20
        
        
        
        'initialize data in variables before starting to loop through
        last_row = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
        open_price = ws.Cells(2, 3).Value
        greatest_percent_increase_value = 0
        greatest_percent_decrease_value = 0
        greatest_total_volume_value = 0
        
                    
        'read in data and loop through each row till last row
        For r = 2 To last_row
        
                            
            current_ticker = ws.Cells(r, 1).Value
            current_volume = CDbl(ws.Cells(r, 7).Value)
            next_ticker = ws.Cells(r + 1, 1).Value
        
            If current_ticker <> next_ticker Then
                'if there is a change between the current ticker and the next ticker write the current ticker to the worksheet
                        
                'write the ticker in the spreadsheet
                ws.Cells(i, 10).Value = current_ticker
                
                'put the yearly change in the spreadsheet and set the background color
                close_price = ws.Cells(r, 6).Value
                yearly_change = close_price - open_price
                
                If yearly_change > 0 Then
                    ws.Cells(i, 11).Interior.ColorIndex = 4
                    'green
                Else
                    'RED
                    ws.Cells(i, 11).Interior.ColorIndex = 3
                End If
                
                ws.Cells(i, 11).Value = yearly_change
                
                'calculate and put the percent change to the spreadsheet
                
                
                If open_price = 0 Then
                    ws.Cells(i, 12).Value = " "
                    'percent_change = 0
                    'percent_change = Format(percent_change, "Percent")
                Else
                    percent_change = (close_price / open_price - 1)
                    'keep the value for calculations
                    percent_change_value = percent_change
                    'format for display
                    percent_change = Format(percent_change, "Percent")
                    ws.Cells(i, 12).Value = percent_change
                End If
                          
                'ws.Cells(i, 12).Value = percent_change
                
                'calculate and display the total stock value to the spreadsheet
                total_volume = total_volume + current_volume
                ws.Cells(i, 13).Value = total_volume
                
                
                'compare current to greatest save if greater
                If percent_change_value > greatest_percent_increase_value Then
                    greatest_percent_increase_ticker = current_ticker
                    greatest_percent_increase_value = percent_change_value
                End If
                
                If percent_change_value < greatest_percent_decrease_value Then
                    greatest_percent_decrease_ticker = current_ticker
                    greatest_percent_decrease_value = percent_change_value
                End If
                
                
                If total_volume > greatest_total_volume_value Then
                    greatest_total_volume_ticker = current_ticker
                    greatest_total_volume_value = total_volume
                End If
                
                
                
                i = i + 1
                'set the total_volume back to zero to start the
                total_volume = 0
                'set the open value to the next open value
                open_price = ws.Cells(r + 1, 3).Value
                
                            
            Else
                'continue to loop through the current ticker information
                'calculate the total stock volume as you loop through
                total_volume = total_volume + current_volume
                
            End If
                
            
        'go on to the next row
        Next r
        
        'do not autofit the columns - it took a really long time
                
        ws.Cells(2, 17).Value = greatest_percent_increase_ticker
        ws.Cells(2, 18).Value = Format(greatest_percent_increase_value, "Percent")
        
     
        ws.Cells(3, 17).Value = greatest_percent_decrease_ticker
        ws.Cells(3, 18).Value = Format(greatest_percent_decrease_value, "Percent")
       
        
        ws.Cells(4, 17).Value = greatest_total_volume_ticker
        ws.Cells(4, 18).Value = greatest_total_volume_value
        
            
    'go on to the next worksheet
    Next
    'Debug.Print current_ticker
    
    
    
End Sub
