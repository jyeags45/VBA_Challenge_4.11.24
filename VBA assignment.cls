Sub Stock():

For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Total Stock Volume"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "% Change"


    ws.Cells(1, 9).EntireColumn.AutoFit
    ws.Cells(1, 10).EntireColumn.AutoFit
    ws.Cells(1, 11).EntireColumn.AutoFit
    ws.Cells(1, 12).EntireColumn.AutoFit


    ws.Range("A1:G1").Interior.Color = RGB(255, 255, 0)
    ws.Range("i1:l1").Interior.Color = RGB(0, 255, 255)


    Dim ticker_name As String
    Dim yr_change As Double
    Dim ticker_vol As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim per_change As Double
    Dim data_summ As Double
    
    
    open_price = ws.Cells(2, 3).Value
    data_summ = 2
    ticker_vol = 0

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For x = 2 To lastrow
        If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
            close_price = ws.Cells(x, 6).Value
            yr_change = (close_price - open_price)
            ticker_name = ws.Cells(x, 1).Value
            ticker_vol = ticker_vol + ws.Cells(x, 7).Value
            ws.Range("I" & data_summ).Value = ticker_name
            ws.Range("j" & data_summ).Value = ticker_vol
            ws.Range("k" & data_summ).Value = yr_change
            
                If open_price = 0 Then
                    per_change = 0
                    
                Else
                    per_change = yr_change / open_price
                    
                End If
            
            ws.Range("l" & data_summ).Value = per_change
            data_summ = data_summ + 1
            ticker_vol = 0
            open_price = ws.Cells(x + 1, 3)
        Else
            ticker_vol = ticker_vol + ws.Cells(x, 7).Value
        End If
        
    Next x
    
   data_summ = ws.Cells(Rows.Count, 9).End(xlUp).Row
   
        For x = 2 To data_summ
         
         If ws.Cells(x, 11).Value > 0 Then
            ws.Cells(x, 11).Interior.Color = RGB(0, 255, 0)
         Else
            ws.Cells(x, 11).Interior.Color = RGB(255, 0, 0)
        End If
        
    Next x
    
  Next ws

End Sub

