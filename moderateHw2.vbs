Sub moderateStockVolume()

' what do I do in each worksheet?
' worksheet parameter and starting point
Dim ws As Worksheet
Dim starting_ws As Worksheet
Sheets(1).Select
Set starting_ws = ActiveSheet

For Each ws In Sheets

ws.Activate

' turn off screen updating to speed up performance
Application.ScreenUpdating = False

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
maxLimit = Cells(Rows.Count, 1).End(xlUp).Row
    
' variable for opening price
Dim openPrice As Variant
        
' variable for closing price
Dim closePrice As Variant
        
        ' variables for yearly change
        Dim yearlyChange As Double
        Dim percentChange As Double
        
            ' other variables
            Dim i As Double
            ' seperate variable for price setting
            Dim finalIteration As String
            
            i = 2
            
            ' dynamic value for iterations
            finalIteration = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' insert headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    ' SET opening price
        openPrice = Cells(i, 3).Value
    
        ' create loop for all stocks
    For i = 2 To finalIteration
        
        ' SET opening stock name
            
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        ' use volume numbers in COLUMN 7 to add to stock total
        stockTotal = stockTotal + ws.Cells(i, 7).Value
            
            Else
            stockName = Cells(i, 1).Value
        ' print stock name in COLUMN 9
        Range("I" & stockRow).Value = stockName
            ' add final volume of stock
            stockTotal = stockTotal + Cells(i, 7).Value
                ' define final closing price
                Cells(i + 1, 6).Select
                ActiveCell.Offset(-1, 0).Select
                closePrice = ActiveCell.Value
                
                ' calculate yearly change
                yearlyChange = closePrice - openPrice
                
                'calculate percent change
                'This might be totally wrong, as I encountered an issue with stock PLNT
                'having a start price of 0 for multiple dates.
                'But under the time constraints, this is the best I could come up with!
                If openPrice <> 0 Then
                percentChange = yearlyChange / openPrice
                Else
                percentChange = 0
                End If
                
                ' reset opening price for next stock
                Cells(i + 1, 3).Select
                openPrice = ActiveCell.Value
                
                        ' print yearly change in COLUMN 10
                        Range("J" & stockRow).Value = yearlyChange
                
                        ' print yearly change in COLUMN 11
                        Range("K" & stockRow).Value = percentChange
                        Range("K" & stockRow).NumberFormat = "0.00%"
            
            ' print stock volume total in COLUMN 12
            Range("L" & stockRow).Value = stockTotal
            
            ' go down to the next available row
            stockRow = stockRow + 1
            
            ' RESET stock total (in order to keep adding to it properly)
            stockTotal = 0
             
            ' use volume numbers in COLUMN 7 to add to stock total
            stockTotal = stockTotal + Cells(i, 7).Value
            
        ' exit if-statement
        End If
        
        ' reset the loop for i
        Next i
                                
                                ' conditional colors
                                For i = 2 To finalIteration
                                If Cells(i, 10).Value > 0 Then
                                Cells(i, 10).Interior.ColorIndex = 4
                                ElseIf Cells(i, 10).Value < 0 Then
                                Cells(i, 10).Interior.ColorIndex = 3
                                'autofit column width for stock totals *in every worksheet*
                                Columns("J").NumberFormat = "0.000000000"
                                Columns("K").AutoFit
                                Columns("L").AutoFit
                                End If
                                Next i
                                
' select next worksheet and continue loop
Next ws

starting_ws.Activate
                                        
' turn screen updating back on
Application.ScreenUpdating = True

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

