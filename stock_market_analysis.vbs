' Aizza Asuncion
' UCBSAN1010Data
' Unit 2 homework
' Dr.Spronck
' 30 October 2017

Sub calculateStockData()
    
  ' Declare and initialize variables
  Dim ticker As String
  
  Dim yearly_change As Single
  Dim year_close_price As Double
  Dim percent_change As Double
  
  Dim year_open_price As Double
  year_open_price = Cells(2, 3).Value
  
  Dim greatest_increase_ticker As String
  Dim greatest_decrease_ticker As String
  Dim greatest_volume_ticker As String
  
  Dim greatest_increase As Double
  greatest_increase = 0
  
  Dim greatest_decrease As Double
  greatest_decrease = 0
  
  Dim greatest_total_volume As Double
  greatest_total_volume = 0

  Dim total_stock_volume As Double
  total_stock_volume = 0

  Dim summary_table_row As Integer
  summary_table_row = 2
  
  Dim i As Long
        
  Dim lastrow As Double
  lastrow = Cells(Rows.Count, "A").End(xlUp).Row
  
  'Create summary table headers
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Volume"
 
    ' Loop through all stocks
    For i = 2 To lastrow
    
        ' Determine if next row has different ticker symbol
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Assign value to ticker
            ticker = Cells(i, 1).Value

            ' Add to the total stock volume
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
            
                ' determine if greatest total volume
                If total_stock_volume > greatest_total_volume Then
                    greatest_total_volume = total_stock_volume
                    greatest_volume_ticker = ticker
                End If
        
            ' Assign value to close price
            year_close_price = Cells(i, 6).Value
            
            ' Calculate yearly change
            yearly_change = year_close_price - year_open_price
        
            ' Calculate percent_change
            percent_change = yearly_change / year_open_price
            
                ' determine if greatest increase
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                    greatest_increase_ticker = ticker
            
                'determine if greatest decrease
                ElseIf percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                    greatest_decrease_ticker = ticker
            
                End If

            ' Print the ticker symbol in the Summary Table
            Range("I" & summary_table_row).Value = ticker
        
            ' Print the yearly change in the Summary Table
            Range("J" & summary_table_row).Value = Round(yearly_change, 9)
            
                ' highlight green if positive change
                If yearly_change >= 0 Then
                     Range("J" & summary_table_row).Interior.ColorIndex = 4
                ' Otherwise highlight red if negative change
                 Else
                    Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
        
            ' Print the percent change in the Summary Table
            Range("K" & summary_table_row).Value = percent_change

            ' Print the total stock volume to the Summary Table
            Range("L" & summary_table_row).Value = total_stock_volume

            ' Add one to the summary table row
            summary_table_row = summary_table_row + 1
      
            ' Reset the total stock volume, yearly change, and percent change
            year_open_price = Cells(i + 1, 3).Value
            total_stock_volume = 0
            yearly_change = 0
            percent_change = 0

        ' If next row has same ticker symbol
        Else
                    
            ' Add to the total stock volume
            total_stock_volume = total_stock_volume + Cells(i, 7).Value

        End If

    Next i

  ' format summary table and greatest info
  Range("J:J").NumberFormat = "0.000000000"
  Range("K:K").NumberFormat = "0.00%"
  Range("Q2:Q3").NumberFormat = "0.00%"
   
  'Create table headers for greatest info
  Range("P1").Value = "Ticker"
  Range("Q1").Value = "Value"

  ' print greatest in another summary table
  Range("O2").Value = "Greatest % Increase"
  Range("P2").Value = greatest_increase_ticker
  Range("Q2").Value = greatest_increase
  Range("O3").Value = "Greatest % Decrease"
  Range("P3").Value = greatest_decrease_ticker
  Range("Q3").Value = greatest_decrease
  Range("O4").Value = "Greatest Total Volume"
  Range("P4").Value = greatest_volume_ticker
  Range("Q4").Value = greatest_total_volume

End Sub