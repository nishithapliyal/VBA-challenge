Attribute VB_Name = "Module1"
Sub stocks()

    For Each ws In ThisWorkbook.Worksheets
        
        ws.Range("I1, P1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim LastRow As Long
        Dim LastRowTicker As Long
        Dim start As Long
        Dim newtickstart As Long
        Dim Inc_percentchange As Double
        Dim Dec_percentchange As Double
        Dim totalstockvol As Double

     
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        start = 2
        newtickstart = 2
        
        
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'copy ticker value
                ws.Cells(start, 9).Value = ws.Cells(i, 1).Value
                
                'calculate quarterly change
                ws.Cells(start, 10).Value = ws.Cells(i, 6).Value - ws.Cells(newtickstart, 3).Value
                
                'conditional formatting of quarterly change col - neg is red and pos is green
                If ws.Cells(start, 10).Value < 0 Then
                    ws.Cells(start, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(start, 10).Interior.ColorIndex = 4
                End If
            
                'calculate percent change and convert to percent format
                ws.Cells(start, 11).Value = ws.Cells(start, 10).Value / ws.Cells(newtickstart, 3).Value
                ws.Cells(start, 11).Value = Format(ws.Cells(start, 11).Value, "Percent")
                'End If
                
                    
                'calculate total stock volume
                ws.Cells(start, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(newtickstart, 7), ws.Cells(i, 7)))
                
                start = start + 1
                newtickstart = i + 1
                
            End If
               
        
        Next i
        
        LastRowTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row
        Inc_percentchange = ws.Cells(2, 11).Value
        Dec_percentchange = ws.Cells(2, 11).Value
        totalstockvol = ws.Cells(2, 12).Value
        
        For i = 2 To LastRowTicker
            
            'find greatest % increase
            If ws.Cells(i, 11).Value > Inc_percentchange Then
                Inc_percentchange = ws.Cells(i, 11).Value
                ws.Cells(2, 17).Value = Inc_percentchange
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            Else
                Inc_percentchange = Inc_percentchange
            End If
            
            'find greatest % decrease
            If ws.Cells(i, 11).Value < Dec_percentchange Then
                Dec_percentchange = ws.Cells(i, 11).Value
                ws.Cells(3, 17).Value = Dec_percentchange
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            Else
                Dec_percentchange = Dec_percentchange
            End If
            
            'find greatest total stock vol
            If ws.Cells(i, 12).Value > totalstockvol Then
                totalstockvol = ws.Cells(i, 12).Value
                ws.Cells(4, 17).Value = totalstockvol
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            Else
                totalstockvol = totalstockvol
            End If
        
            
            ws.Cells(2, 17).Value = Format(ws.Cells(2, 17).Value, "Percent")
            ws.Cells(3, 17).Value = Format(ws.Cells(3, 17).Value, "Percent")
        
        Next i
                
    ws.Columns("A:R").AutoFit
    
    Next ws


End Sub

'quarterly change = last close - first open
'percent change = 100* quarterly change/first open
'total stock vol = sum

        



