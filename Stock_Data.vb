Sub Ticker_()
For Each ws In Worksheets

  ' Set an initial variable for holding the Ticker name'
  Dim Ticker As String

  ' Set an initial variable for holding the total per Ticker'
  Dim Volume_Total As Double
  Volume_Total = 0

  ' Keep track of the location for Ticker in the summary table'
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  
  'find last row of cells with values'
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  'Set initial variable for yearly Change'
  Dim Ticker_Open As Double
  Dim Ticker_Close As Double
  Dim Yearly_Change As Double
  
  
  Ticker_Open = ws.Range("C2").Value
  Ticker_Close = 0
  Yearly_Change = 0
  
  
  
  'Row Labels'
  
  ws.Cells(1, 10).Value = "Ticker"
  ws.Cells(1, 11).Value = "Volume Total"
  ws.Cells(1, 12).Value = "Yearly Change"
  ws.Cells(1, 13).Value = "Percent Change"
  

  ' Loop through all Ticker Totals'
  For i = 2 To LastRow


' checks to make sure cells contain non-zeros'

    If ws.Cells(i, 3).Value = 0 Then
    Exit For

    End If

    ' Check if we are still within the same Ticker'
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    

      ' Set the Ticker name
       Ticker = ws.Cells(i, 1).Value

      ' Add to the Volume Total'
      Volume_Total = Volume_Total + Cells(i, 7).Value
      
      '*find close price per Ticker'
      Ticker_Close = ws.Cells(i, 6).Value

      ' Print the Ticker in the Summary Table'
      ws.Range("J" & Summary_Table_Row).Value = Ticker

      ' *Print the Ticker Amount to the Summary Table'
      ws.Range("K" & Summary_Table_Row).Value = Volume_Total

      
      ' *Calculate yearly change'
      Yearly_Change = Ticker_Close - Ticker_Open
      
      '*Print Yearly change to Summary Table'
      
      ws.Range("L" & Summary_Table_Row).Value = Yearly_Change
      
      'Calculate percent change and print it to Summary_Table_Row'
      
      ws.Range("M" & Summary_Table_Row).Value = (Ticker_Close - Ticker_Open) / Ticker_Open
    
      
      ' *find open yearly price per Ticker not in A'
    
      Ticker_Open = ws.Cells(i + 1, 3).Value
    



      ' Add one to the summary table row'
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total'
      Volume_Total = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Brand Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

    End If

  Next i
  
LastSrow = ws.Cells(Rows.Count, 13).End(xlUp).Row



        For c = 2 To LastSrow

        If ws.Cells(c, 13).Value < 0 Then
    
         ws.Cells(c, 13).Interior.ColorIndex = 3
     
         Else
    
         ws.Cells(c, 13).Interior.ColorIndex = 4
         
         End If
         
        
 Next c
        
  
Next ws

End Sub
