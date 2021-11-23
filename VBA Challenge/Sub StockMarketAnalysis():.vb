Sub StockMarketAnalysis():
    Dim total_vol As Double
    Dim ticker As String
    Dim ticker_counter As double 
    Dim ticker_value As Double
    Dim open_year As double 
    Dim end_year As Double

    For Each ws In Worksheets
        total_vol = 0
        ticker_counter = 2        
        ticker_value = 2   

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To LastRow
            total_vol = total_vol + ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            open_year = ws.Cells(ticker_value, 3)

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                end_year = ws.Cells(i, 6)
                ws.Cells(ticker_counter, 9).Value = ticker
                ws.Cells(ticker_counter, 10).Value = end_year - open_year
                

                If open_year = 0 Then
                    ws.Cells(ticker_counter, 11).Value = Null
                Else
                    ws.Cells(ticker_counter, 11).Value = (end_year - open_year) / open_year
                End If
                ws.Cells(ticker_counter, 12).Value = total_vol

                ' Color the cell green if > 0, red if < 0
                If ws.Cells(ticker_counter, 10).Value > 0 Then
                    ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3
                End If

                
                total_vol = 0
                ticker_counter = ticker_counter + 1
                ticker_value = i + 1 

            End If

        Next i

    Next ws

End Sub

