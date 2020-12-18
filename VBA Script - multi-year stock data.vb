VBA-challenge-script

Sub Stock()

'Create a for loop to cycle through worksheets
For Each ws In Worksheets

    'Set an initial variables
    Dim worksheetname As String
    worksheetname = ws.Name
    Dim Ticker_Name As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Volume As LongLong
    Dim PreviousRowClose As Double
    Dim Open_Rate As Double
    Open_Rate = ws.Cells(2, 3).Value
    x = 2
    Volume = 0

    'Create locations for the output data
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"

    'find endrow
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Last_Summary_Row = ws.Cells(Rows.Count, 10).End(xlUp).Row

    'create loop for rows
    For i = 2 To Last_Row
      
      PreviousRowClose = ws.Cells(i, 6).Value
         
       'Check to see if the next row's ticker is not equal to current row's ticker
       If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
          'Set close rate for cells
          Close_Rate = ws.Cells(i, 6).Value
                 
          'Set variable for ticker name (ROW, column 1)
          TickerName = ws.Cells(i, 1).Value
          
          'Print ticker name into summary table
          ws.Range("I" & x).Value = TickerName
                              
          'Calculate volume
          ws.Range("L" & x).Value = Volume + ws.Cells(i, 7).Value
          
          'Calculate yearly change
          Yearly_Change = Close_Rate - Open_Rate
          ws.Range("J" & x).Value = Round(Yearly_Change, 2)
          
            'Calculate % change
            If Open_Rate = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = (Yearly_Change / Open_Rate) * 100
                ws.Range("K" & x).Value = "%" & Percent_Change
            
            End If

          'Add one to summary table row
          x = x + 1
          
          'Reset values
          Volume = 0
          Yearly_Change = 0
          Percent_Change = 0
          Open_Rate = ws.Cells(i + 1, 3).Value
           
          'If the next cell matches the current cell,
          Else
          
             'Calculate values within matching cells
             Volume = Volume + ws.Cells(i, 7).Value
             ws.Range("J" & x).Value = Yearly_Change
          
        End If
                                                     
    Next i
    
    For i = 2 To Last_Summary_Row
    
    'Conditional formatting for yearly change
        If ws.Cells(i, 10).Value >= 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        
        'Otherwise turn them...
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        
        End If
    
    Next i
               
 Next ws

End Sub