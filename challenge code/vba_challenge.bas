Attribute VB_Name = "vba_challenge"
Sub looping()


Dim ws As Worksheet


For Each ws In ThisWorkbook.Sheets


    Dim lastrow As Long
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
    Dim uniTicker As Integer
        uniTicker = 2
        
    Dim totalvolume As Double
        totalvolume = 0
        
    Dim uniticker_volume As Double
        uniticker_volume = 12
        
    Dim volumesummarycol As Integer
        volumesummarycol = 12
    
    Dim firstopencol As Double
        firstopencol = 20
    
    Dim lastclosecol As Double
        lastclosecol = 21
    
    Dim firstopen As Double
        firstopen = ws.Cells(2, 3)
        'starts here because loop skips the first row, starts at difference
    
    Dim lastclose As Double
        lastclose = 0
    
    Dim yearlychange As Double
        yearlychange = 0
    
    Dim percentchange As Double
        percentchange = 0
    
    Dim greatestincrease As Double
        greatestincrease = 0
    
    Dim unigreatestincrease As Double

    Dim greatestdecrease As Double
        greatestdecrease = 0
    
    Dim unigreatestdecrease As Double
      
    Dim greatestincreaseticker As String
    
    Dim greatestdecreaseticker As String
    
    Dim unigreatestvolume As Double
    
    Dim greatestvolume As Double
        greatestvolume = 0
    
    Dim greatestvolumeticker As String
    
    
    'define ranges for each title
    
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    
    ws.Range("p1").Value = "Ticker"
    ws.Range("q1").Value = "Value"
    ws.Range("o2").Value = "Greatest % Increase"
    ws.Range("o3").Value = "Greatest % Decrease"
    ws.Range("o4").Value = "Greatest Total Volume"
    
    ws.Range("i1:l1").EntireColumn.AutoFit
    ws.Range("o2:o3").EntireColumn.AutoFit
    ws.Range("k:k").NumberFormat = "0.00%"
    ws.Range("q2:q3").NumberFormat = "0.00%"
    
    'this loop retrieves and records each stocks unique ticker, their total stock volume,
    ' and their first open/last close
        For i = 2 To lastrow
        
            totalvolume = totalvolume + ws.Cells(i, 7)
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
                ws.Cells(uniTicker, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(uniTicker, volumesummarycol).Value = totalvolume
                lastclose = ws.Cells(i, 6).Value
                yearlychange = lastclose - firstopen
                ws.Cells(uniTicker, 10).Value = yearlychange
                percentchange = (yearlychange / firstopen)
                ws.Cells(uniTicker, 11).Value = percentchange
                firstopen = ws.Cells(i + 1, 3).Value
            
                yearlychange = 0
                percentchange = 0
                totalvolume = 0
                uniTicker = uniTicker + 1
            
        
            End If
            Next i
        
    'color cells
    For i = 2 To lastrow
    
        If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value <= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
        
        
        End If
        Next i
    
    'greatest % increase value + ticker
    For i = 2 To lastrow
    
        unigreatestincrease = ws.Cells(i, 11).Value
        If unigreatestincrease > greatestincrease Then
        greatestincrease = unigreatestincrease
        greatestincreaseticker = ws.Cells(i, 9).Value
        ws.Range("q2") = greatestincrease
        ws.Range("p2") = greatestincreaseticker
        
        End If
        Next i
    
    'greatest % decrease value + ticker
    For i = 2 To lastrow
    
        unigreatestdecrease = ws.Cells(i, 11).Value
        If unigreatestdecrease < greatestdecrease Then
        greatestdecrease = unigreatestdecrease
        greatestdecreaseticker = ws.Cells(i, 9).Value
        ws.Range("q3") = greatestdecrease
        ws.Range("p3") = greatestdecreaseticker
        
        End If
        Next i
    
    'greatest total volume value + ticker
    For i = 2 To lastrow
        
        unigreatestvolume = ws.Cells(i, 12).Value
        If unigreatestvolume > greatestvolume Then
        greatestvolume = unigreatestvolume
        greatestvolumeticker = ws.Cells(i, 9).Value
        ws.Range("q4") = greatestvolume
        ws.Range("p4") = greatestvolumeticker
        
        End If
        Next i
    

Next ws




End Sub


        
