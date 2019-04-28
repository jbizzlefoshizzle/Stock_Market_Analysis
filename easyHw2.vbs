Sub easyStockVolume()

Dim ws As Object

' what do I do in each worksheet?
For Each ws In Worksheets

    ' variable for stock name
    Dim stockName As String
    
    ' variable for stock total volume
    Dim stockTotal As Double
    stockTotal = 0
    
    'track location for each stock volume summary
    Dim stockRow As Integer
    stockRow = 2
    
    ' set dynamic limit for each worksheet
    Dim maxLimit As String
    maxLimit = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' insert headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    
        ' create loop for all stocks
        For i = 2 To maxLimit
            
            ' if the next cell is of a different stock
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            ' set stock name
            stockName = ws.Cells(i, 1).Value
            
            ' use volume numbers in COLUMN 7 to add to stock total
            stockTotal = stockTotal + ws.Cells(i, 7).Value
            
            ' print stock name in COLUMN 9
            ws.Range("I" & stockRow).Value = stockName
            
            ' print stock volume total in COLUMN 10
            ws.Range("J" & stockRow).Value = stockTotal
            
            ' go down to the next available row
            stockRow = stockRow + 1
            
            ' RESET stock total (in order to keep adding to it properly)
            stockTotal = 0
            
        ' if the next cell is of the same stock
        Else
            
            ' use volume numbers in COLUMN 7 to add to stock total
            stockTotal = stockTotal + ws.Cells(i, 7).Value
            
                    ' exit if-statement
                    End If
        
                        ' reset the loop
                        Next i
                                'autofit column width for stock totals *in every worksheet*
                                ws.Columns("J").AutoFit
                                

' go to next worksheet
Next ws
            
End Sub

Sub resetStockSummary()

' what do I do in each worksheet?
For Each ws In Worksheets

    ' select COLUMN 9 AND COLUMN 10 and *clear data*
    ws.Range("I:I").Clear
    ws.Range("J:J").Clear
    ws.Range("K:K").Clear
    ws.Range("L:L").Clear

        ' select next worksheet and continue loop
        Next ws
        
End Sub
