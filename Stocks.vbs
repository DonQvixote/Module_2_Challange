Sub Stocks():

    Dim i, n, k As Long
    Dim j, x As Integer
    Dim ticker, iticker, dticker, vticker As String
    Dim op, cl, yc, pc, vol, gi, gd, gvol As Double
    Dim sheet(0 To 2) As String
        sheet(0) = "2018"
        sheet(1) = "2019"
        sheet(2) = "2020"
        
    For x = 0 To 2
        
        Sheets(sheet(x)).Activate
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest %Increase"
        Range("O3").Value = "Greatest %Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("p1").Value = "Ticker"
        Range("q1").Value = "Value"
        Range("q2:q3").NumberFormat = "0.00%"
        
        
        n = Worksheets(sheet(x)).UsedRange.Rows.Count
        k = 2
        j = 2
        
        For i = 2 To n
            ticker = Cells(i, 1).Value
            op = Cells(k, 3).Value
            vol = Cells(i, 7).Value + vol
            
            If ticker <> Cells(i + 1, 1).Value Then
                cl = Cells(i, 6).Value
                yc = cl - op
                pc = yc / op
                
                Cells(j, 9).Value = ticker
                Cells(j, 10).Value = yc
                Cells(j, 11).Value = pc
                Cells(j, 11).NumberFormat = "0.00%"
                Cells(j, 12).Value = vol
                
                If yc > 0 Then
                    Cells(j, 10).Interior.ColorIndex = 4
                Else
                    Cells(j, 10).Interior.ColorIndex = 3
                End If
                
                If pc > 0 Then
                    Cells(j, 11).Interior.ColorIndex = 4
                Else
                    Cells(j, 11).Interior.ColorIndex = 3
                End If
                
                If gi < pc Then
                    gi = pc
                    iticker = ticker
                ElseIf pc < gd Then
                    gd = pc
                    dticker = ticker
                End If
                
                If gvol < vol Then
                    gvol = vol
                    vticker = ticker
                End If
                        
            vol = 0
            j = j + 1
            k = i + 1
                
            End If
        
        Next i
            
        Range("q2").Value = gi
        Range("P2").Value = iticker
        Range("q3").Value = gd
        Range("P3").Value = dticker
        Range("q4").Value = gvol
        Range("P4").Value = vticker
                    
        gi = 0
        gd = 0
        gvol = 0
        
    Next x
    
End Sub