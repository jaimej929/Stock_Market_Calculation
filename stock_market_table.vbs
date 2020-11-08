
Sub Stock_market_calculations()


For Each ws In Worksheets
'Headers____________________________
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Year Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    
'Dims___________________________________
  Dim ticker_name As String
  Dim year_change As Double
  Dim percent_change As Double
  Dim Total_sv As Double
  Dim col As Double
  Dim High As Double
  Dim low As Double
  Dim Ticker_count As Double
  Dim start_tick As Double
  Dim close_tick As Double
  
  'Set values________________
  Ticker_count = 0
  close_tick = 0
  start_tick = 0
  Total_sv = 0
  
  
  
  Dim table As Integer
  table = 2
  
  col = ws.Cells(Rows.Count, "A").End(xlUp).Row
  For i = 2 To col
    
    
    'start if statements____________________
    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        Ticker_count = Ticker_count + 1
    End If
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker_name = ws.Cells(i, 1).Value
        
        
        
        
        
        Total_sv = Total_sv + ws.Cells(i, 7).Value
        
        close_tick = ws.Cells(i, 6).Value
        
        open_tick = ws.Cells(i - Ticker_count, 3).Value
        
        year_change = close_tick - open_tick
        
        percent_change = (close_tick - open_tick) / open_tick

    ws.Range("k" & table).Value = percent_change
    
    ws.Columns("K:K").NumberFormat = "0.00%"
    
    ws.Range("j" & table).Value = year_change
    
    ws.Columns("J:J").NumberFormat = "$0.00"
        
    ws.Range("i" & table).Value = ticker_name
        
    ws.Range("l" & table).Value = Total_sv
        
            If year_change > 0 Then
        
        ws.Range("j" & table).Interior.ColorIndex = 4
    Else
       
       ws.Range("j" & table).Interior.ColorIndex = 3
    
    End If
        
        
        
        table = table + 1
        
        Total_sv = 0
        
        percent_change = 0
         
        
    

    Else
        Total_sv = Total_sv + Cells(i, 7).Value
              
    End If
    

              
        
    Next i
   
  Next ws

End Sub
