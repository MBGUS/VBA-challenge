Sub stockmarket()

    For Each ws In Worksheets
    
    'Set an initial variable for holding the Ticker name
        Dim Ticker_Name As String
    
    'Set an initial variable for holding the total per volumen
        Dim Volume_Total As Double
        Volume_Total = 0
    
    'Set an initial variable for holding yearly change (open - close value)
        Dim Yearly_Change As Double
        Yearly_Change = 0
        
    'Set an initial variable for holding percent change (open - close value)
        Dim Percent_Change As Double
        Dim Start_value As Double
        Percent_Change = 0
        Start_value = 2
    
    'Keep track of the location for each tikcer in the summary table (starts in row 2 where the data beggins)
        Dim Summary_Table As Double
        Summary_Table = 2
        
    'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Loop through all stock market data
    For i = 2 To LastRow
    
        'Check if we are still within the same ticker, if we are not (Ticker A != Ticker A)...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Set the Ticker name (on the summary)
            Ticker_Name = ws.Cells(i, 1).Value
            
            'Add the Volume total (operation SUM Volume)
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            
            'Add the Yaerly Change value (operation Start Open - Last Close)
            Yearly_Change = (Yearly_Change + (ws.Cells(i, 6).Value - ws.Cells(Start_value, 3).Value))
            
            'Add the Percent Change value --- (if division / 0 = 0 then operation: Close-Open/Open)
             If ws.Cells(Start_value, 3).Value = 0 Then
                Percent_Change = 0
                
                Else
                Percent_Change = (Yearly_Change / ws.Cells(Start_value, 3).Value)
                
                'Formating
                ws.Range("K:K").NumberFormat = "0.00%"
                                       
                End If
                       
            'Print the Ticker name in the Summary Table (result)
            ws.Range("I" & Summary_Table).Value = Ticker_Name
            
            'Print the Volume total in the Summary Table (result)
            ws.Range("L" & Summary_Table).Value = Volume_Total
    
            'Print the Yearly Change in the Summary Table (result)
            ws.Range("J" & Summary_Table).Value = Yearly_Change
            
            'Add color to cells (>0 green / <0 red)
            If Yearly_Change > 0 Then
            
                ws.Range("J" & Summary_Table).Interior.ColorIndex = 4
            
                Else
         
                ws.Range("J" & Summary_Table).Interior.ColorIndex = 3
            
            End If
            
            'Print the Percent Change in the Summary Table (result)
            ws.Range("K" & Summary_Table).Value = Percent_Change
            
            'Add one to the summary table row
            Summary_Table = Summary_Table + 1
            
            'Reset the Volume Total
            Volume_Total = 0
            
            'Reset the Yearly Change
            Yearly_Change = 0
            
            'Reset the Change
            Percent_Change = 0
            
            'Reset the Start value (***)
            Start_value = i + 1
               
        'If cell inmediately following a row is the same brand... (Ticker A = Ticker A)
        
        Else
        
            'Add to the Total Volume (for the next round)
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
        
        End If
        
    Next i
   
'Set variables for summary table
max_inc = WorksheetFunction.Max(ws.Range("K:K"))
max_inc_tic = WorksheetFunction.Match(max_inc, ws.Range("K:K"), 0)
max_dic = WorksheetFunction.Min(ws.Range("K:K"))
max_dic_tic = WorksheetFunction.Match(max_dic, ws.Range("K:K"), 0)
max_vol = WorksheetFunction.Max(ws.Range("L:L"))
max_vol_tic = WorksheetFunction.Match(max_vol, ws.Range("L:L"), 0)

'Bring back the result for each variable
ws.Range("Q2").Value = max_inc
ws.Range("P2").Value = ws.Cells(max_inc_tic, 9).Value
ws.Range("Q3").Value = max_dic
ws.Range("P3").Value = ws.Cells(max_dic_tic, 9).Value
ws.Range("Q4").Value = max_vol
ws.Range("P4").Value = ws.Cells(max_vol_tic, 9).Value
    
'Formating
ws.Range("Q2:Q3").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = Number
    
    Next ws
    
End Sub