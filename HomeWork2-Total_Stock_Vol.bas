Attribute VB_Name = "Module1"

Sub StockData()
    For Each ws In Worksheets
        ws.Activate
        
        Range("I1").value = "Ticker"
        Range("J1").value = "Yearly Change"
        Range("K1").value = "Percent Change"
        Range("L1").value = "Total Stock Volume"
        Range("O2").value = "Greatest Percent Increase"
        Range("O3").value = "Greatest Percent Decrease"
        Range("O4").value = "Greatest Total Stock Volume"
        Range("P1").value = "Ticker"
        Range("Q1").value = "Value"
   
        Columns("A:Q").AutoFit
        
        Dim ticker As String, total_stock_volume As LongLong
        Dim yearly_change As Double, percent_change As Double
        Dim dateRanges As Range, Latest As String, Earliest As String
   
        total_stock_volume = 0
   
        ticker_row = 2
        opening_price = Cells(2, "C").value
        row_count = Cells(Rows.Count, "A").End(xlUp).Row
        
        For Row = 2 To row_count
        
            If Cells(Row + 1, "A").value <> Cells(Row, "A").value Then
               ticker = Cells(Row, "A").value
               total_stock_volume = total_stock_volume + Cells(Row, "G").value
               closing_price = Cells(Row, "F").value
               
               yearly_change = closing_price - opening_price
               
               percentage_change = (yearly_change / opening_price) * 100
               
               Range("I" & ticker_row).value = ticker
               Range("J" & ticker_row).value = yearly_change
               Range("K" & ticker_row).value = "%" & percentage_change
               Range("L" & ticker_row).value = total_stock_volume
               
                If yearly_change < 0 Then
        
                    Cells(ticker_row, 10).Interior.ColorIndex = 3
            
                Else
            
                    Cells(ticker_row, 10).Interior.ColorIndex = 4
                
                End If
               
               ticker_row = ticker_row + 1
               total_stock_volume = 0
               
               opening_price = Cells(Row + 1, "C").value
               
            Else
            
                total_stock_volume = total_stock_volume + Cells(Row, "G").value
                
            End If
        
        Next Row
        
        greatest_increase = 0
        
        row_count = Cells(Rows.Count, "I").End(xlUp).Row

        For Row = 2 To row_count
        
            If Cells(Row, "K").value > greatest_increase Then
            
                greatest_increase = Cells(Row, "K").value
                ticker = Cells(Row, "I")
                
            End If
            
        Next Row
        
            ws.Range("P2").value = ticker
            ws.Range("Q2").value = greatest_increase
            
            greatest_decrease = 0
        
        row_count = Cells(Rows.Count, "I").End(xlUp).Row

        For Row = 2 To row_count
        
            If Cells(Row, "K").value < greatest_decrease Then
            
                greatest_decrease = Cells(Row, "K").value
                ticker = Cells(Row, "I")
                
            End If
            
        Next Row
        
            ws.Range("P3").value = ticker
            ws.Range("Q3").value = greatest_decrease
            
            greatest_total_stock_vol = 0
        
        row_count = Cells(Rows.Count, "L").End(xlUp).Row

        For Row = 2 To row_count
        
            If Cells(Row, "L").value > greatest_total_stock_vol Then
            
                greatest_total_stock_vol = Cells(Row, "L").value
                ticker = Cells(Row, "I")
                
            End If
            
        Next Row
        
            Range("P4").value = ticker
            Range("Q4").value = greatest_total_stock_vol
                
    Next ws
    
    MsgBox ("complete")
         
End Sub
