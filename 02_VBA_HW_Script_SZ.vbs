Sub Stock()

   For Each ws In Worksheets
    
   'Set an intial variable for holding ticker
   Dim Ticker As String
   
   'Set an initial variable for holing the Total Stock Volume
   Dim Total_Stock_Volume As Double
   Total_Stock_Volume = 0
   
   'Keep track of the location for each stock in the summary table
   Dim Summary_Table_Row As Long
   Summary_Table_Row = 2
   
   'Set an intiial variable for Opening_Price, Closing_Price, Previous_Price
   Dim Opening_Price As Double
   Dim Closing_Price As Double
   Dim Previous_Price As Double
   Previous_Price = 2

  'Set an intial variable for holding the Yearly Change
   Dim Yearly_Change As Double
   
   'Set an initial variable for holding the Percent Change
   Dim Percent_Change As Double
   Percent_Change = 0
   
   'Set LastRow as variable
   Dim LastRow As Long
   
   'Set initial variables for Greatest_%_Increase, Greatest_%_Decrease, and Greatest_Total_Volume
   Dim Greatest_Increase As Double
   Greatest_Increase = 0
   Dim Greatest_Decrease As Double
   Greatest_Decrease = 0
   Dim Greatest_Total_Volume As Double
   Greatest_Total_Volume = 0
   
   'Add headers
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly_Change"
   ws.Cells(1, 11).Value = "Percent_Change"
   ws.Cells(1, 12).Value = "Total_Stock_Volume"
   ws.Cells(1, 16).Value = "Ticker"
   ws.Cells(1, 17).Value = "Value"
   ws.Cells(2, 15).Value = "Greatest_%_Increase"
   ws.Cells(3, 15).Value = "Greatest_%_Decrease"
   ws.Cells(4, 15).Value = "Greatest_Total_Volume"
   
   
   'Determine the last Row
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
   'Loop through all stock
   For i = 2 To LastRow
   
    'Add to Total Stock Volume
   Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
   
   'Check if we are still within the same stock ticker
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
   'Set Ticker Name
   Ticker = ws.Cells(i, 1).Value
   
   'Print the ticker info in the summary table
   ws.Range("I" & Summary_Table_Row).Value = Ticker
   
   'Print the Total Stock Volume Amount in the Summary table
   ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
   
   'Reset Total_Stock_Volume
   Total_Stock_Volume = 0
   
    'Set Opening_Price, Closing_Price, Yearly_Change
   Opening_Price = ws.Range("C" & Previous_Price)
   Closing_Price = ws.Range("F" & i)
   
   Yearly_Change = Closing_Price - Opening_Price
   
   'Print the Yearly Change in the Summary table
   ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
   
   'Calculate Percent Change
   If Opening_Price <> 0 Then
    Opening_Price = ws.Range("C" & Previous_Price)
    Percent_Change = (Closing_Price - Opening_Price) / Opening_Price
    
    End If
   
    'Print the Percent Change in the Summary table
   ws.Range("K" & Summary_Table_Row).Value = Percent_Change
   ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
   
   'Reset Percent_Change
   Percent_Change = 0
   
   Summary_Table_Row = Summary_Table_Row + 1
   Previous_Price = i + 1
   
   End If
   
   'Conditional formatting for Yearly_Change
   If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    Else
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        
    End If
   
Next i
   
   'Greatest_%_Increase, Great_%_Decrease, Greatest_Total_Volume
   
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
   For i = 2 To LastRow
   
   If ws.Range("K" & i).Value > ws.Cells(2, 17).Value Then
    ws.Cells(2, 17).Value = ws.Range("K" & i).Value
    ws.Cells(2, 16).Value = ws.Range("I" & i).Value
    
    End If
    
    If ws.Range("K" & i).Value < ws.Cells(3, 17).Value Then
    ws.Cells(3, 17).Value = ws.Range("K" & i).Value
    ws.Cells(3, 16).Value = ws.Range("I" & i).Value
    
    End If
    
    If ws.Range("L" & i).Value > ws.Cells(4, 17).Value Then
    ws.Cells(4, 17).Value = ws.Range("L" & i).Value
    ws.Cells(4, 16).Value = ws.Range("I" & i).Value
    
    End If
    
    Next i
    
    'Format %
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
   Next ws
    
End Sub




