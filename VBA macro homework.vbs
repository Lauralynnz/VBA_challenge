Sub wallstreet_ticker()

'----- Loopthrough sheets ----
    Dim ws As Worksheet

    For Each ws In Worksheets

'----- Create New Columns ----

          ' Set an initial variable for holding the ticker symbol
          Dim Ticker As String
          
          ' Set an initial variable for holding the yearly change
          Dim Yearly_Change As Double
          
          ' Set an initial variable for holding the percent change
          Dim Percent_Change As Double
          
          ' Set an initial variable for holding total stock volume
          Dim Stock_Total As Double
          Stock_Total = 0
          
          ' Define value of open column
          Open_Value = ws.Cells(2, 3).Value
          
          ' Keep track of the location for each ticker in the summary table
          Dim Summary_Table_Row As Long
          Summary_Table_Row = 2
          
            ' Create column headers
            ws.Cells(1, 10).Value = "Ticker"
                ws.Cells(1, 10).Font.Bold = "True"
            ws.Cells(1, 11).Value = "Yearly Change"
                ws.Cells(1, 11).Font.Bold = "True"
            ws.Cells(1, 12).Value = "Percent Change"
                ws.Cells(1, 12).Font.Bold = "True"
            ws.Cells(1, 13).Value = "Stock Total"
                ws.Cells(1, 13).Font.Bold = "True"
            
            ' Autosize the new columns
            ws.Columns("J:M").EntireColumn.AutoFit
        
        
        '----- Populate New Columns ----
          
          ' Loop through all tickers using last_row
          LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
          ' MsgBox (LastRow)
                    For i = 2 To LastRow
        
                ' Check if we are still within the same ticker, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Calculate the yearly  and percent change
                Close_Value = ws.Cells(i, "F").Value
                        Yearly_Change = Close_Value - Open_Value
                        Percent_Change = (Yearly_Change / Open_Value) * 100
                        ws.Range("L" & Summary_Table_Row).Value = Percent_Change
                        
                Open_Value = ws.Cells(i + 1, "C").Value
                    ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
            
            'color positives in green and negatives in red
                      If Percent_Change > 0 Then
                            ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                        ElseIf Percent_Change < 0 Then
                            ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                        Else
                            ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 0
                End If
            
                  ' Set the ticker name
                Ticker = ws.Cells(i, 1).Value
            
                  ' Add to the stock total
               Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            
                  ' Print the ticker name in the Summary Table
                 ws.Range("J" & Summary_Table_Row).Value = Ticker
            
                  ' Print the total stock volume to the Summary Table
                ws.Range("M" & Summary_Table_Row).Value = Stock_Total
            
                  ' Add one to the summary table row
                 Summary_Table_Row = Summary_Table_Row + 1
                  
                  ' Reset the Brand Total
                 Stock_Total = 0
            
                ' If the cell immediately following a row is the same ticker...
               Else
            
                  ' Add to the Stock Total
                 Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            
               End If
        
          Next i
  
  Next ws

End Sub



