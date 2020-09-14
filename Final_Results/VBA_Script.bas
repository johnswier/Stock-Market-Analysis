Attribute VB_Name = "Module1"
Sub LoopScript():
    For Each ws In Worksheets
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Total Stock Volume"
        ws.Range("K1") = "Yearly Change"
        ws.Range("L1") = "Percent Change"
        ws.Range("O1") = "Value"
        last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        Dim ticker As String
        Dim yearly_change As Double
        Dim first_open As Double
        Dim last_close As Double
        Dim counter As Long
        
        counter = 1
        For Row = 2 To last_row
            ticker = ws.Cells(Row, 1)
            
            If ticker = ws.Cells(Row - 1, 1) Then
                ws.Cells(counter, 10) = ws.Cells(counter, 10) + ws.Cells(Row, 7)
            
                If ws.Cells(Row + 1, 1) <> ticker Then
                    last_close = ws.Cells(Row, 6)
                    yearly_change = last_close - first_open
                    ws.Cells(counter, 11) = yearly_change
                    
                    If first_open = 0 Then
                        ws.Cells(counter, 12) = 0
                    Else
                        percentage_change = yearly_change / first_open
                        ws.Cells(counter, 12) = percentage_change
                        ws.Cells(counter, 12).NumberFormat = "0.00%"
                    End If
                    
                End If
            
            Else
                counter = counter + 1
                first_open = ws.Cells(Row, 3)
                
                ws.Cells(counter, 9) = ticker
                ws.Cells(counter, 10) = ws.Cells(Row, 7)
            End If
        Next Row
        
        ws.Cells(2, 14) = "Greatest % Increase"
        ws.Cells(3, 14) = "Greatest % Decrease"
        ws.Cells(4, 14) = "Greatest Total Stock Volume"
        
        ws.Cells(2, 15) = Application.WorksheetFunction.Max(ws.Range("L:L"))
        ws.Cells(3, 15) = Application.WorksheetFunction.Min(ws.Range("L:L"))
        ws.Cells(4, 15) = Application.WorksheetFunction.Max(ws.Range("J:J"))
        
        ws.Cells(2, 15).NumberFormat = "0.00%"
        ws.Cells(3, 15).NumberFormat = "0.00%"
        
        lrow_yearly_change = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
        For Row = 2 To lrow_yearly_change
            ws.Cells(Row, 12) = FormatPercent((ws.Cells(Row, 12)), 2)
            If ws.Cells(Row, 11) >= 0 Then
                ws.Cells(Row, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(Row, 11) < 0 Then
                ws.Cells(Row, 11).Interior.ColorIndex = 3
            End If
        Next Row
    Next ws
End Sub
