Sub tickerSummary()

  'Set an initial variable for holding the stock price change
  Dim yearlyPriceChange As Double
  yearlyPriceChange = 0
        
  'Set an initial variable for holding the percentage change
  Dim percentageChange As Double
  percentageChange = 0
  
  'Set an initial variable for holding total stock volume
  Dim totalVolume As Double
  totalVolume = 0
  
  'Keep track of the location to populate the summary
  Dim summaryTable As Integer
  summaryTable = 2
  Range("I1").Value = "Ticker"
  Range("J1").Value = "PriceChange"
  Range("K1").Value = "PercentageChange"
  Range("L1").Value = "TotalStockVolume"
        
  'Set a variable to track the starting row for each ticker; when the ticker changes the index is reassigned
  Dim indexTracker As Long
  indexTracker = 2
        
  'Set last row
  lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
    'Set loop for all rows
    For i = 2 To lastRow
    
        'Set beginning date as minimum value and ending date as maximum value
        Dim beginDate As Long
        Dim endDate As Long
                
        beginDate = WorksheetFunction.Min(Cells(i, 2).Value)
        endDate = WorksheetFunction.Max(Cells(i, 2).Value)
        
        'Set open price and closing price
        Dim openPrice As Double
        Dim closingPrice As Double
            
        openPrice = Cells(indexTracker, 3).Value
        closingPrice = Cells(i, 6).Value
        
        'Set a variable to track the current ticker
        Dim currentTicker As String
        currentTicker = Cells(i, 1)
        
        'Set a variable to track when ticker changes
        Dim nextTicker As String
        nextTicker = Cells(i + 1, 1)
                
          If currentTicker <> nextTicker And openPrice = 0 Then
                        
            'Log the current ticker to the summary table
              Range("I" & summaryTable).Value = currentTicker
              
              'Log the price changeto the summary table
              Range("J" & summaryTable).Value = 0
                                    
          'Color code green if price change is positive and red if negative
          If yearlyPriceChange = 0 Then
              Range("J" & summaryTable).Interior.ColorIndex = 0
              
          End If
                                    
              'Log the percent change in the summary table
              Range("K" & summaryTable).Value = Format(0, "Percent")

              'Log total volume in the summary table
              Range("L" & summaryTable).Value = 0
              
              'Add one to summary table row
              summaryTable = summaryTable + 1
              
              'reset yearly change
              yearlyPriceChange = 0
              
              'Set the new starting index for different ticker
              indexTracker = i + 1
              
              'Reset total volume
              totalVolume = 0
                        
          'Check if the next value is the same as the current value, if not then
          ElseIf currentTicker <> nextTicker And openPrice <> 0 Then
                        
              'Take the closing price at the ending date and subtract it by open price at begining date
              yearlyPriceChange = closingPrice - openPrice
              
              'calculate percentage change
              percentageChange = (yearlyPriceChange / openPrice)
              
              'sum volume to totalVolume
              totalVolume = totalVolume + Cells(i, 7).Value
                                
              'Log the current ticker to the summary table
              Range("I" & summaryTable).Value = currentTicker
              
              'Log the price changeto the summary table
              Range("J" & summaryTable).Value = yearlyPriceChange
                                    
              'Color code green if price change is positive and red if negative
              If yearlyPriceChange < 0 Then
                  Range("J" & summaryTable).Interior.ColorIndex = 3
              
              Else
                  Range("J" & summaryTable).Interior.ColorIndex = 4
                  
              End If
                                    
              'Log the percent change in the summary table
              Range("K" & summaryTable).Value = Format(percentageChange, "0.00%")
  
              'Log total volume in the summary table
              Range("L" & summaryTable).Value = totalVolume
              
              'Add one to summary table row
              summaryTable = summaryTable + 1
              
              'reset yearly change
              yearlyPriceChange = 0
              
              'Set the new starting index for different ticker
              indexTracker = i + 1
              
              'Reset total volume
              totalVolume = 0
  
              Else
              
                  'Add volume to total
                  totalVolume = totalVolume + Cells(i, 7).Value
                            
                                
              End If
                    
    Next i
     
        'Challenge of the assignment
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
     

         Dim greatestPercentIncrease As Double
         Dim greatestPercentDecrease As Double
         Dim greatestTotalVolume As Double
         
         greatestPercentIncrease = 0
        
              For j = 2 To lastRow
                     
                If Cells(j, 11).Value > greatestPercentIncrease Then
                  greatestPercentIncrease = Cells(j, 11).Value
                  Range("P2").Value = Cells(j, 9).Value
                  Range("Q2").Value = Cells(j, 11).Value
                  Range("Q2").NumberFormat = "0.00%"
                
                End If
                        
              Next j
                
                        
              For k = 2 To lastRow
            
                If Cells(k, 11).Value < greatestPercentDecrease Then
                  greatestPercentDecrease = Cells(k, 11).Value
                  Range("P3").Value = Cells(k, 9).Value
                  Range("Q3").Value = Cells(k, 11).Value
                  Range("Q3").NumberFormat = "0.00%"
                        
                End If
                    
              Next k
                        
              For m = 2 To lastRow
            
                If Cells(m, 12).Value > greatestTotalVolume Then
                  greatestTotalVolume = Cells(m, 12).Value
                  Range("P4").Value = Cells(m, 9).Value
                  Range("Q4").Value = Cells(m, 12).Value
                           
                End If
                        
              Next m

'Setting columns to best fit
Range("A:Z").Columns.AutoFit
     
End Sub
