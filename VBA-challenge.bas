Attribute VB_Name = "Module1"
Sub stock_calc()
    
    For Each ws In Worksheets
    
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
    
        Dim curr_ticker As String
        Dim curr_row As Long
        Dim curr_open As Double
        Dim curr_close As Double
        Dim curr_total As Double
        Dim i As Long
    
        curr_ticker = ws.Cells(2, 1)
        curr_row = 2
        curr_open = 0
        curr_close = 0
        curr_total = 0
        i = 2
        
        While StrComp(curr_ticker, "") <> 0
    
            curr_open = ws.Cells(i, 3)
            curr_total = 0
        
            While StrComp(curr_ticker, ws.Cells(i, 1)) = 0
                curr_total = curr_total + ws.Cells(i, 7)
                curr_close = ws.Cells(i, 6)
                i = i + 1
            
            Wend
        
            ws.Cells(curr_row, 9) = curr_ticker
            ws.Cells(curr_row, 10) = curr_close - curr_open
            ws.Cells(curr_row, 11) = ws.Cells(curr_row, 10) / curr_open
            ws.Cells(curr_row, 12) = curr_total
            curr_row = curr_row + 1
            curr_ticker = ws.Cells(i, 1)
        
        Wend
        
        i = 2
    
        While StrComp(ws.Cells(i, 9), "") <> 0
            If ws.Cells(i, 10) < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
                ws.Cells(i, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 4
                ws.Cells(i, 11).Interior.ColorIndex = 4
            End If
            i = i + 1
        Wend
        
        Dim max_inc As Long
        Dim max_dec As Long
        Dim max_tot As Long
    
        i = 3
        max_inc = 2
        max_dec = 2
        max_tot = 2
    
        While StrComp(ws.Cells(i, 9), "") <> 0
            If ws.Cells(max_inc, 11) < ws.Cells(i, 11) Then
                max_inc = i
            End If
        
            If ws.Cells(max_dec, 11) > ws.Cells(i, 11) Then
                max_dec = i
            End If
        
            If ws.Cells(max_tot, 12) < ws.Cells(i, 12) Then
                max_tot = i
            End If
        
            i = i + 1
        Wend
    
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
    
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
    
        ws.Range("Q2") = ws.Cells(max_inc, 1)
        ws.Range("Q3") = ws.Cells(max_dec, 1)
        ws.Range("Q4") = ws.Cells(max_tot, 1)
    
        ws.Range("Q2") = ws.Cells(max_inc, 11)
        ws.Range("Q3") = ws.Cells(max_dec, 11)
        ws.Range("Q4") = ws.Cells(max_tot, 12)
        
    Next ws
    
End Sub
