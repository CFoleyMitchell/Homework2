Attribute VB_Name = "Module1"
Sub Ticker_Total_EntireWrkbk()
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
For Each ws In Worksheets

  ' Set an initial variable for holding the ticker name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total volume per ticker
  Dim Ticker_Volume_Total As Double
  Ticker_Volume_Total = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all ticker records
   For i = 2 To LastRow

    ' Check if we are still within the same ticker name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Ticker Volume Total
      Ticker_Volume_Total = Ticker_Volume_Total + ws.Cells(i, 7).Value

      ' Print the Ticker Name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Ticker Volume Amount to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Ticker_Volume_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      Ticker_Volume_Total = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Ticker Total
      Ticker_Volume_Total = Ticker_Volume_Total + ws.Cells(i, 7).Value

    End If

  Next i
  
  'Label and Format the Summary_Table_Row
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Total Stock Volume"
  ws.Columns("I:J").EntireColumn.AutoFit
  
Next ws

End Sub
