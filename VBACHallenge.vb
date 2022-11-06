Sub VBACHallenge()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

    'Assign column headings

    ws.Cells(1, 11).Value = "Ticker Name"
    ws.Cells(1, 12).Value = "Yearly Change"
    ws.Cells(1, 13).Value = "Percentage Change"
    ws.Cells(1, 14).Value = "Volume of stock"
    ws.Cells(2, 17).Value = "Greatest % Increase"
    ws.Cells(3, 17).Value = "Greatest % Decrease"
    ws.Cells(4, 17).Value = "Greatest Total Volume"
    ws.Cells(1, 18).Value = "Ticker Name"
    ws.Cells(1, 19).Value = "Value"


    'Assign Variables for the Activity
    
        'Assign Ticker name varia
       Dim Ticker_Name As String
    
      ' Set an initial variable for holding the total
        Dim Ticker_Total As Double
        Ticker_Total = 0
    
      ' Keep track of the location for each ticker
        Dim Ticker_Row As Double
        Ticker_Row = 2
    
      'Assign Variable for opening price and closing price
      Dim Closing_Price As Double
      Dim Opening_Price As Double
      Dim Yearly_Change As Double
      Dim Percentage_Change As Double
      Opening_Price = ws.Cells(2, 3).Value
      
     ' Assign variables for Bonus
      Dim Greatest_Perc_Increase As Double
      Dim Greatest_Perc_Decrease As Double
      Dim VolTicker As Double
      
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
           
    ' Loop through all tickers
      For i = 2 To LastRow
      
        ' Check if we are still within the ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
          ' Set the Ticker
          Ticker_Name = ws.Cells(i, 1).Value
    
          ' Add to the Ticker Total
          Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
         
          'Assign closing price
          Closing_Price = ws.Cells(i, 6).Value
          
          'Calculating yearly change
           Yearly_Change = Closing_Price - Opening_Price
           
           'Percentage Change
           Percentage_Change = (Yearly_Change / Opening_Price)
            
            
       ' Print Outputs
       
          ' Print the Ticker
          ws.Range("K" & Ticker_Row).Value = Ticker_Name
          
          ' Print the Yearly Change
          ws.Range("L" & Ticker_Row).Value = Yearly_Change
          
          ' Print the Percentage Change
          ws.Range("M" & Ticker_Row).Value = Percentage_Change
          ws.Range("M" & Ticker_Row).NumberFormat = "0.00%"
    
          ' Print the Ticker Total
          ws.Range("N" & Ticker_Row).Value = Ticker_Total
              
          
          ' Conditional Formatting
           If Yearly_Change < 0 Then
             ws.Range("L" & Ticker_Row).Interior.ColorIndex = 3
          
           Else
             ws.Range("L" & Ticker_Row).Interior.ColorIndex = 4
            
          End If
                  
          ' Add one to the row
          Ticker_Row = Ticker_Row + 1
          
          'Assign opening price for next ticker
          Opening_Price = ws.Cells(i + 1, 3).Value
          
          ' Reset the Ticker Total
          Ticker_Total = 0
    
        ' If the cell immediately following a row is the same ticker
        
        Else
    
          ' Add to the Ticker Total
          Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

    End If

  Next i

''''''''''''''''''''''''''''''''''''''''''''''''''''

' Bonus part

         Dim GreatestTicker As String
         Dim SmallestTicker As String
         Dim GreatestVol As String
        
         LastRow2 = ws.Cells(Rows.Count, 13).End(xlUp).Row
        
     For i = 2 To LastRow2

         If Greatest_Perc_Increase < ws.Cells(i, 13).Value Then
           Greatest_Perc_Increase = ws.Cells(i, 13).Value
           GreatestTicker = ws.Cells(i, 11).Value
           
         End If
         If Greatest_Perc_Decrease > ws.Cells(i, 13).Value Then
         Greatest_Perc_Decrease = ws.Cells(i, 13).Value
         SmallestTicker = ws.Cells(i, 11).Value
         End If
        If VolTicker < ws.Cells(i, 14).Value Then
           VolTicker = ws.Cells(i, 14).Value
           GreatestVol = ws.Cells(i, 11).Value
        End If
    Next i
    
    ws.Cells(2, 19).Value = Greatest_Perc_Increase
    ws.Cells(2, 18).Value = GreatestTicker
    ws.Cells(3, 19).Value = Greatest_Perc_Decrease
    ws.Cells(3, 18).Value = SmallestTicker
    ws.Cells(4, 19).Value = VolTicker
    ws.Cells(4, 18).Value = GreatestVol

    ws.Range("S2:S3").NumberFormat = "0.00%"
    
Next ws

End Sub
