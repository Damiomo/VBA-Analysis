Sub Stock():
    dim ws as worksheet

    for each ws in worksheets

        ' Headers
        ws.[I1] = "Ticker"
        ws.[J1] = "Quaterly Change"
        ws.[K1] = "Percentage Change"
        ws.[L1] = "Total Stock Volume"
        ws.[P1] = "Ticker"
        ws.[Q1] = "Value"
        ws.[O2] = "Greatest%increase"
        ws.[O3] = "Great%decrease"
        ws.[O4] = "Greatest Total Volume"
        
        ' Variables
        si = 2
        Total = 0
        firstOpen = 0
        greatestVol = 0
        greatestIncPct = 0
        greatestDecPct = 0
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To lastRow
        ' sum of all volume for the ticker
            Total = Total + ws.Cells(i, "G")
        
        ' Store first open of ticker
            If firstOpen = 0 Then
                firstOpen = ws.Cells(i, "C")
            End If
                
        ' Last row of ticker
            If ws.Cells(i, "a") <> ws.Cells(i + 1, "a") Then
                ws.Cells(si, "I") = ws.Cells(i, "A")
            
            ' Quaterly Change
                    qCh = ws.Cells(i, "F") - firstOpen
                    ws.Cells(si, "J") = qCh
                
            ' Color green for number above zero
                    If qCh > 0 Then
                        ws.Cells(si, "J").Interior.ColorIndex = 4
                    End If
                    
            ' Color red for number below zero
                    If qCh < 0 Then
                        ws.Cells(si, "J").Interior.ColorIndex = 3
                    End If
                    
            ' Percentage Change
                    perCh = qCh / firstOpen
                    ws.Cells(si, "K") = perCh
                    
            ' Total Stock Volume
                    ws.Cells(si, "L") = Total

            ' Greatest Increase Percentage
                If perCh > greatestIncPct Then
                    greatestIncPct = perCh
                    greatestIncTicker = ws.Cells(i, "A")
                End If

            ' Greatest Decrease Percentage
                If perCh < greatestDecPct Then
                    greatestDecPct = perCh
                    greatestDecTicker = ws.Cells(i, "A")
                End If

            ' Greatest Volume
                If Total > greatestVol Then
                    greatestVol = Total
                    greatestTicker = ws.Cells(i, "A")
                End If
                
            ' Resets for next ticker
                Total = 0
                si = si + 1
                firstOpen = 0
            End If
            
        Next i

        ' Populate Greatest Increase, Decrease and Volume
            ws.Range("P2") = greatestIncTicker
            ws.Range("P3") = greatestDecTicker
            ws.Range("P4") = greatestTicker

            ws.Range("Q2") = greatestIncPct
            ws.Range("Q3") = greatestDecPct
            ws.Range("Q4") = greatestVol
        
        ' Table Formatting
            ws.Columns.AutoFit
            ws.Columns("K").NumberFormat = "##.##%"
            ws.Columns("Q").NumberFormat = "##.##%"

next ws
End Sub