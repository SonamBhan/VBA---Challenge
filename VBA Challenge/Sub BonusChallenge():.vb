Sub BonusChallenge():

        Dim greatest_inc As double 
        Dim greatest_dec As Double
        Dim greatest_dec_row As interger
        Dim greatest_inc_row As integer 
        Dim GreatestTotalVol_index As Integer
        Dim GreatestTotalVol As Double

    For Each ws In Worksheets
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"


        greatest_inc = 0
        greatest_dec = 0
        GreatestTotalVol = 0
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        For i = 2 To LastRow
           
            If ws.Cells(i, 11) > greatest_inc Then
                greatest_inc = ws.Cells(i, 11)
                greatest_inc_row = i
            End If

            If ws.Cells(i, 11) < greatest_dec Then
                greatest_dec = ws.Cells(i, 11)
                greatest_dec_row = i
            End If

            
            If ws.Cells(i, 12) > GreatestTotalVol Then
                GreatestTotalVol = ws.Cells(i, 12)
                GreatestTotalVol_index = i
            End If

        Next i

        ws.Range("P2") = ws.Cells(greatest_inc_row, 9).Value
        ws.Range("P3") = ws.Cells(greatest_dec_row, 9).Value
        ws.Range("P4") = ws.Cells(GreatestTotalVol_index, 9).Value

        ws.Range("Q2") = greatest_inc
        ws.Range("Q3") = greatest_dec
        ws.Range("Q4") = GreatestTotalVol


    Next ws

End Sub
