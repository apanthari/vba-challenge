Sub module_2()


For Each ws In Worksheets

    

    'intiate vars
    Dim i As Long
    Dim start As Long
    Dim change As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim percentChange As Double
    Dim totalstock As Double
    Dim grtincre As Double
    Dim grtincretick As String
    Dim grtdec As Double
    Dim grtdectick As String
    Dim greatestVolume As Double
    Dim greatestVolumetick As String
    
    Dim ticker As String
    
    
    
    'intiate first value of counter
    start = 2
    
    'intiate the start value for open
    openPrice = ws.Cells(2, "C").Value
    
    'intiate stock value
    totalstock = 0
    greatestVolume = 0
    
    'iniate greatest increase
    grtincre = 0
    grtdec = 0
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    'start for loop that will go through ALL rows
    For i = 2 To RowCount
    
        'check if the current value is not the same as the next value
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
            
        
            'ticker value
            ticker = ws.Cells(i, 1).Value
            ws.Cells(start, "I").Value = ticker
            
            
            'change close-open
            closePrice = ws.Cells(i, "F").Value
            change = closePrice - openPrice
            ws.Cells(start, "J").Value = change
            
            'percentage change
            percentChange = (closePrice - openPrice) / openPrice
            ws.Cells(start, "K").Value = percentChange

            
            
            'total stock price
            totalstock = Cells(i, "G").Value + totalstock
            ws.Cells(start, "L").Value = totalstock
        
            
            If percentChange > grtincre Then
                grtincre = percentChange
                grtincretick = ticker
        
            ElseIf percentChange < grtdec Then
                grtdec = percentChange
                grtdectick = ticker
            
            ElseIf totalstock > greatestVolume Then
                greatestVolume = totalstock
                greatestVolumetick = ticker
            
            End If
            
            
            totalstock = 0
            
            'greatest percentage increase
    
            
            'go to the next row in the printed values
            start = start + 1
            
            'reset open price to the new value which is the next cell value
            openPrice = ws.Cells(i + 1, "C").Value

            
        ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            totalstock = ws.Cells(i, "G").Value + totalstock
            
        End If
     
        
        'color the yearly change values
    
    
        If ws.Cells(i, "J").Value > 0 Then
            ws.Cells(i, "J").Interior.ColorIndex = 43
        Else
            ws.Cells(i, "J").Interior.ColorIndex = 3
        End If
        
        
    Next i
    
    
    'print greatest/lowest increase value
    ws.Cells(2, "P").Value = grtincre
    ws.Cells(2, "O").Value = grtincretick
    
    ws.Cells(3, "P").Value = grtdec
    ws.Cells(3, "O").Value = grtdectick
        
    ws.Cells(4, "P").Value = greatestVolume
    ws.Cells(4, "O").Value = greatestVolumetick
    
    'format headers
    
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "L").Value = "Total Stock Volume"
    ws.Cells(1, "O").Value = "Ticker"
    ws.Cells(1, "P").Value = "Value"
    ws.Cells(2, "N").Value = "Greatest % increase"
    ws.Cells(3, "N").Value = "Greatest % decrease"
    ws.Cells(4, "N").Value = "Greatest Stock Increase"
    


    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("P2:P3").NumberFormat = "0.00%"
    
Next ws

End Sub
