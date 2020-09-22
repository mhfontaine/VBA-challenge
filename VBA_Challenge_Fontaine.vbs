
Sub VBAHomework()
    For Each ws In Worksheets
        ws.Activate
        Call CalculateSummary
    Next ws
End Sub

'The code follows is a script that loops through several stocks for one year.
'The script, when successfully executed will display:
'    - The ticker symbol.
'    - Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'    - The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'    - The total stock volume of the stock.

' Code here starts here
Sub CalculateSummary()
    
'Declarations
Dim ticker As String
Dim next_ticker As String
Dim total As Integer
Dim totalrows As Double
Dim summary_row As Integer
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
'Dim percent_change As Long
Dim percent_change2 As Double
Dim total_stock_volume As Double
     
percent_change = CLng(2000) * 10
    
open_price = Cells(2, 3).Value
         
'Finding the number of total rows in the spreadsheet
totalrows = Cells(Rows.Count, "A").End(xlUp).Row
summary_row = 2
    
'This area of code sets the title rows on the spreadsheet
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
    
'This area of code is for the challenge area only.  Challenge not attempted.
 Range("O2").Value = "Greatest % Increase"
 Range("O3").Value = "Greatest % Decrease"
 Range("O4").Value = "Greatest Total Volume"
 Range("I:O").Columns.AutoFit
    
    
        For currentrow = 2 To totalrows
                          
                'Print ticker symbol
                ticker = Cells(currentrow, 1).Value
                next_ticker = Cells(currentrow + 1, 1).Value
                prev_ticker = Cells(currentrow - 1, 1).Value
                total_stock_volumn = 0
                Range("I" & summary_row).Value = open_price
        
                'Yearly change
                closing_price = Cells(currentrow, 6).Value
                yearly_change = closing_price - open_price
                Range("J" & summary_row).Value = yearly_change
    
        
        If ticker = next_ticker Then
                total_stock_volumn = total_stock_volumn + Cells(currentrow, 7).Value
                If ticker <> prev_ticker Then
                    open_price = Cells(currentrow, 3).Value
                End If
        
          Else
                total_stock_volumn = total_stock_volumn + Cells(currentrow, 7).Value
                close_price = Cells(currentrow, 6).Value
                percent_change2 = (close_price - open_price)
                percent_change = (percent_change2 / open_price) * 100
               
        
                Cells(summary_row, 9).Value = ticker
                Cells(summary_row, 10).Value = yearly_change
                Cells(summary_row, 11).Value = percent_change
                Cells(summary_row, 12).Value = total_stock_volumn
                Cells(summary_row, 11).NumberFormat = "0.00"
            
               If Cells(summary_row, 10).Value > 0 Then
                     Cells(summary_row, 10).Interior.ColorIndex = 4
               ElseIf Cells(summary_row, 10).Value < 0 Then
                     Cells(summary_row, 10).Interior.ColorIndex = 3
               End If
            
              summary_row = summary_row + 1
              total_stock_volumn = 0
              percent_change = 0
              yearly_change = 0
              open_price = 0
              close_price = 0
           
        End If
        
        Next currentrow
       
    Debug.Print ActiveSheet.Name
    
End Sub


