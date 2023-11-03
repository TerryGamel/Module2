Attribute VB_Name = "Module1"
Sub CollateData():
    Dim ws As Worksheet
    For Each ws In Worksheets
    
        Dim GISymbol, GDSymbol, GTSymbol As String
        Dim GIValue, GDValue, GTValue As Double
    
        GIValue = 0
        GDValue = 0
        GTValue = 0
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        Dim LastRow As Long
        LastRow = FindLastRow(ws)
    
        Dim TickerCounter As Long
        Dim StartPrice, EndPrice, Volume As Double
    
        TickerCounter = 2
        StartPrice = 0
        EndPrice = 0
        Volume = 0
    
        For i = 2 To LastRow
            Volume = Volume + ws.Cells(i, 7).Value
            If StartPrice = 0 Then
                StartPrice = ws.Cells(i, 3).Value
            End If
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ws.Cells(TickerCounter, 9).Value = ws.Cells(i, 1).Value
                EndPrice = ws.Cells(i, 6).Value
                
                Dim PriceDif, PercentDif As Double
                PriceDif = EndPrice - StartPrice
                PercentDif = PriceDif / StartPrice
                
                ws.Cells(TickerCounter, 10).Value = PriceDif
                ws.Cells(TickerCounter, 11).Value = PercentDif
                ws.Cells(TickerCounter, 11).NumberFormat = "0.00%"
                ws.Cells(TickerCounter, 12).Value = Volume
                
                If PriceDif < 0 Then
                    ws.Cells(TickerCounter, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(TickerCounter, 10).Interior.ColorIndex = 4
                End If
                If PercentDif > GIValue Then
                    GIValue = PercentDif
                    GISymbol = ws.Cells(i, 1).Value
                ElseIf PercentDif < GDValue Then
                    GDValue = PercentDif
                    GDSymbol = ws.Cells(i, 1).Value
                End If
                If Volume > GTValue Then
                    GTValue = Volume
                    GTSymbol = ws.Cells(i, 1).Value
                End If
                
                Volume = 0
                StartPrice = 0
                EndPrice = 0
                TickerCounter = TickerCounter + 1
            End If
        Next i
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("P2").Value = GISymbol
        ws.Range("Q2").Value = GIValue
        ws.Range("P3").Value = GDSymbol
        ws.Range("Q3").Value = GDValue
        ws.Range("P4").Value = GTSymbol
        ws.Range("Q4").Value = GTValue
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Columns("A:Q").AutoFit
    Next ws

End Sub

Function FindLastRow(ws As Worksheet) As Long
    FindLastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
End Function

