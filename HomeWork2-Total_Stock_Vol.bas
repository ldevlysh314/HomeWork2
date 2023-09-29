Attribute VB_Name = "Module1"
Sub StockData()

'Totals For Stock Volumes
    'set the variables
    Dim ticker As String, total_stock_volume As LongLong
   
    total_stock_volume = 0
   
     'location of a ticker
      ticker_row = 2
       'identify the range of rows
        For Row = 2 To 753001
            'conditions
            If Cells(Row + 1, 1).Value <> Cells(Row, 1) Then
               ticker = Cells(Row, 1).Value
               total_stock_volume = total_stock_volume + Cells(Row, 7).Value
               Range("I" & ticker_row).Value = ticker
               Range("L" & ticker_row).Value = total_stock_volume
               ticker_row = ticker_row + 1
               total_stock_volume = 0
            Else
                total_stock_volume = total_stock_volume + Cells(Row, 7).Value
            End If
        Next Row
    
    
End Sub

Sub YearlyChange()

'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'set the variables

Dim ticker As String, yearly_change As Double
Dim dateRanges As Range, Latest As String, Earliest As String
Set dateRanges = Range("b2:b753001")

With Application
Latest = .Match(.Max(dateRanges), dateRanges, 0)
End With

With Application
Earliest = .Match(.Min(dateRanges), dateRanges, 0)
End With

Dim close_value As Double, open_value As Double
      
      ticker_row = 2
      
      For Row = 2 To 753001
      
     If Cells(Row, 1).Value = "AAB" And Cells(Row, 2).Value = Latest Then
        close_value = Cells(Row, 6)
     End If
     
     If Cells(Row, 1).Value = "AAB" And Cells(Row, 2).Value = Earliest Then
        open_value = Cells(Row, 3)
     End If
    
    Next Row
    
        Range("J" & ticker_row).Value = yearly_change
        Range("I" & ticker_row).Value = ticker

      
yearly_change = close_value - opden_value.Value
         
End Sub
