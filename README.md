# Stocks-Analysis

## Stock Analysis for steve's parents
   I would  like a recommend on my  basis on Analysis .I think They  wiil invest on ENPH and Run stocks  because They wii get good return.
   
   
## Challenge
   I got Challenge for Analysis, I need  perform the same analysis of stocks in 2018, but  i need  a switch the nesting order of your for loops in order to run the analysis faster. After Analysis,Code is look like blow
   
   Sub AllStockAnalysisRefactor()

   '(1) Format  for worksheet
   
        yearValue = InputBox("What year would you like to run the analysis on?")
        Worksheets("All Stocks Analysis Refactor").Activate
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        Cells(3, 1).Value = "Year"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

   '(2) Array of all tickers
   
        Dim tickers(11) As String
   
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
   
   '(3)Arrays for starting price and ending price
   
       Dim startingPrice(11) As Double
       Dim endingPrice(11) As Double
       Dim totalVolume(11) As Double
       Dim tickerIndex As Integer
   
       tickerIndex = 0
   
      'set all indexes to zero
       tickerIndex = 0
       For i = 0 To 11
       startingPrice(i) = 0
       endingPrice(i) = 0
       totalVolume(i) = 0
        
       Next i
    
   '(4)loop through row
       Worksheets(yearValue).Activate
   
       RowCount = Cells(Rows.Count, "A").End(xlUp).Row

       For i = 2 To RowCount
       
       For j = 0 To 11
             tickerIndex = j
             
             If Cells(i, 1).Value = tickers(tickerIndex) Then
             totalVolume(tickerIndex) = totalVolume(tickerIndex) + Cells(i, 8).Value
             End If
                        
             If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
             startingPrice(tickerIndex) = startingPrice(tickerIndex) + Cells(i, 6).Value
             End If
            
             If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
             endingPrice(tickerIndex) = endingPrice(tickerIndex) + Cells(i, 6).Value
                    
             End If
             Next j
             Next i
       

       
   
       'Formatting for Worksheets
        Worksheets("All Stocks Analysis Refactor").Activate
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns("B").AutoFit
        Columns("A").AutoFit
        
    '(5)output data
        For i = 0 To 11
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = totalVolume(i)
        Cells(4 + i, 3).Value = endingPrice(i) / startingPrice(i) - 1

        'Conditional formating
         If Cells(4 + i, 3) > 0 Then
         Cells(4 + i, 3).Interior.Color = vbRed
         ElseIf Cells(4 + i, 3) < 0 Then
         Cells(4 + i, 3).Interior.Color = vbGreen
         Else
         Cells(4 + i, 3).Interior.Color = xlNone

        End If

       Next i
   
End Sub



Rewriting code is like more efficient and save time and frustration.As per rewriting code ENPH and Run has good growth that why I recommend to Steve'parents to invest in this stocks.
