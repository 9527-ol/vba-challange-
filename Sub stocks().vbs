Sub stocks()

        Dim ws As Worksheet
        For Each ws In Worksheets
    
        Dim ticker As String
        
        Dim yearlychange As Double
        yearlychange = 0
        
        Dim percentagechange As Double
        percentagechange = 0
        
        Dim totalstockvolume As Double
        totalstockvolume = 0
        
        Dim summary As Integer
        summary = 2
        
        Dim lastrow As Double
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 7).Value = ticker
        ws.Cells(1, 8).Value = yearlychange
        ws.Cells(1, 9).Value = percentagechange
        ws.Cells(1, 10).Value = totalstockvolume
        
        
        For i = 2 To lastrow
        
        openingprice = ws.Cells(i, 3).Value
        If openingprice = 0 Then
            openingprice = 1
        End If
      
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        
        yearlychange = yearlychange + (ws.Cells(i, 6).Value - openingprice)
        percentagechange = percentagechange + ((ws.Cells(i, 6).Value - openingprice) / openingprice)
        totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
                
                summary = summary + 1
                
                yearlychange = 0
                percentagechange = 0
                totalstockvolume = 0
                
                Else
                
                yearlychange = yearlychange + (ws.Cells(i, 6).Value - openingprice)
                percentagechange = percentagechange + (ws.Cells(i, 6).Value - openingprice) / openingprice
                totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
                
                End If
            If (ws.Cells(summary, 9).Value > 0) Then
            ws.Cells(summary, 9).Interior.ColorIndex = 4
            Else
            ws.Cells(summary, 9).Interior.ColorIndex = 3
            End If
        
        Next i
        
    Next ws
        
End Sub

