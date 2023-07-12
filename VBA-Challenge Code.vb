Sub StockDataAnalysis():

    ' Define all the variables
    
    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalStockVolume As Double
    Dim LastRow As Long
    Dim SummaryTableRow As Integer
    Dim TickerWithTheGreatestIncrease As String
    Dim TickerWithTheGreatestDecrease As String
    Dim TickerWithTheGreatestTotalVolume As String
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0
        
    ' --------------------------------------------------------------------------
    ' Loop through all sheets
    ' --------------------------------------------------------------------------
    
    For Each ws In ThisWorkbook.Worksheets
       LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    ' Create headers for the summary table and "additional functionality table"
       ws.Range("I1").Value = "Ticker"
       ws.Range("J1").Value = "Yearly Change"
       ws.Range("K1").Value = "Percent Change"
       ws.Range("L1").Value = "Total Volume"
       ws.Range("R1").Value = "Ticker"
       ws.Range("S1").Value = "Value"
       ws.Range("O2").Value = "Greatest % Increase"
       ws.Range("O3").Value = "Greatest % Decrease"
       ws.Range("O4").Value = "Greatest Total Volume"
    
    ' Set value of Volume and Summary Table
    
       TotalStockVolume = 0
       SummaryTableRow = 2
        
    ' --------------------------------------------------------------------------
    ' Loop through all rows of data
    ' Get the ticker symbol
    ' Get the opening price for each ticker
    ' Get the closing price for each ticker
    ' Get the total volume for each ticker
    ' Get the yearly change for each ticker
    ' Get the percentage change for each ticker
    ' --------------------------------------------------------------------------
    
        For i = 2 To LastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                OpeningPrice = ws.Cells(i, 3).Value
            End If
            
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
              
            'Check the next cell to make sure the ticker symbol hasn't changed. If no change, continue the previous iteration.
            'Once it has changed, proceed with the calculations
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ClosingPrice = ws.Cells(i, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                PercentageChange = YearlyChange / OpeningPrice
                                
                ' Put results in Summary Table
                 
                ws.Cells(SummaryTableRow, 9).Value = Ticker
                ws.Cells(SummaryTableRow, 10).Value = YearlyChange
                ws.Cells(SummaryTableRow, 11).Value = PercentageChange
                ws.Cells(SummaryTableRow, 12).Value = TotalStockVolume
               
               ' Format PercentageChange Column to a Percent format
                ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
                
               ' Conditional formatting to add green and red cells based on positive or negative values
               ' Leave zero values with White Cells
               
                If YearlyChange > 0 Then
                    ws.Cells(SummaryTableRow, 10).Interior.Color = vbGreen
                ElseIf YearlyChange < 0 Then
                    ws.Cells(SummaryTableRow, 10).Interior.Color = vbRed
                Else
                    ws.Cells(SummaryTableRow, 10).Interior.Color = vbWhite
                End If
                              
                  
           ' ------------------------------------------------------------
           ' Check for greatest increase
           ' Greatest decrease
           ' Greatest volume
           ' ------------------------------------------------------------
            
                If PercentageChange > GreatestIncrease Then
                    GreatestIncrease = PercentageChange
                    TickerWithTheGreatestIncrease = Ticker
                End If
                
                If PercentageChange < GreatestDecrease Then
                    GreatestDecrease = PercentageChange
                    TickerWithTheGreatestDecrease = Ticker
                End If
                
                If TotalStockVolume > GreatestVolume Then
                    GreatestVolume = TotalStockVolume
                    TickerWithTheGreatestVolume = Ticker
                End If
            
    
           ' Put results in a new table
           ' Create headers for the Greatest Increase, Greatest Decrease, and Greatest Volume
            
                
                ws.Cells(2, 18).Value = TickerWithTheGreatestIncrease
                ws.Cells(2, 19).Value = GreatestIncrease
                ws.Cells(2, 19).NumberFormat = "0.00%"
                ws.Cells(3, 18).Value = TickerWithTheGreatestDecrease
                ws.Cells(3, 19).Value = GreatestDecrease
                ws.Cells(3, 19).NumberFormat = "0.00%"
                ws.Cells(4, 18).Value = TickerWithTheGreatestVolume
                ws.Cells(4, 19).Value = GreatestVolume
            
    
            ' Reset the variable for the next ticker
                TotalVolume = 0
                SummaryTableRow = SummaryTableRow + 1
                
            End If
            
        Next i
        
    Next ws
    
End Sub
