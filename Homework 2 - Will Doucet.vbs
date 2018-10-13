Sub Homework()

    Dim LineCount As Long
    Dim CurrentTicker As String
    Dim rng
    Dim TotalVolume
    Dim UniqueTicker As Long
    Dim OldCount As Long
    Dim StartPrice As Double
    Dim EndPrice As Double
    Dim AbsoluteChange As Double
    Dim PercentageChange As Double
    
    
    For Each ws In Worksheets

        'Set/Reset Count For Each Worksheet
        LineCount = 0
        OldCount = 0
        UniqueTicker = 1
        


        'Take First Ticker Symbol In Worksheet
        CurrentTicker = ws.Cells(LineCount + 2, 1).Value

        'Get Last Row in Current Worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Format Column K to Show %
        ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
        
        'Format Column J to Four Decimal Points
        ws.Range("J2:J" & LastRow).NumberFormat = "0.00"
        
        'Format Stock Price Cells
        ws.Range("C2:F" & LastRow).NumberFormat = "0.00"
        
        'Determine the Number of Unique Ticker Symbols
        For i = 2 To LastRow

            'Checks to See If The Current Cell is has a Different Ticker and the Next Cell Down isn't Empty
            If ws.Cells(i, 1).Value <> CurrentTicker And IsEmpty(ws.Cells(i + 1, 1)) = False Then
                'Prints Current Ticker
                ws.Cells(UniqueTicker + 1, "I").Value = CurrentTicker
                'Adds One to UniqueTicker Count
                UniqueTicker = UniqueTicker + 1
                'Takes New Ticker Symbol
                CurrentTicker = ws.Cells(i, 1).Value

            End If

        Next i
        
        'Adds Last Ticker
        ws.Cells(UniqueTicker + 1, "I").Value = CurrentTicker

        'Loop Through Range for Each Unique Ticker Symbol
        For i = 1 To UniqueTicker

            'Update Counts and Reset Total Volume
            OldCount = LineCount
            CurrentTicker = ws.Cells(LineCount + 2, 1).Value
            TotalVolume = 0
            
            'Checks if Current Cell has Same Ticker
            For j = 2 To LastRow

                If ws.Cells(j, 1).Value = CurrentTicker Then

                    LineCount = LineCount + 1

                End If

            Next j
            
            'Calculates Required Info
            rng = ws.Range("G" & (OldCount + 2) & ":G" & (LineCount + 1))
            TotalVolume = Excel.WorksheetFunction.Sum(rng)
            
            StartPrice = ws.Range("C" & (OldCount + 2)).Value
            
            EndPrice = ws.Range("F" & (LineCount + 1)).Value
            
            AbsoluteChange = EndPrice - StartPrice
                        
            If StartPrice > 0 Then
            
                PercentageChange = (AbsoluteChange / StartPrice)
                
                ws.Cells(i + 1, "K").Value = PercentageChange
            
            Else
            
                ws.Cells(i + 1, "K").Value = "N/A"
                
            End If
            
            
            ws.Cells(i + 1, "J").Value = AbsoluteChange

            ws.Cells(i + 1, "L").Value = TotalVolume

        Next i
        
        'Make Column Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Columns("A:L").AutoFit

    Next ws

End Sub

Sub Extra()
    
    Dim FormatRange As Range
    Dim cell As Range
    Dim MaxValue As Double
    Dim MinValue As Double
    Dim MinTicker As String
    Dim MaxTicker As String
    Dim MaxVolume
    Dim VolTicker As String
    Dim LastTicker As Long
    
For Each ws In Worksheets

    LastTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row

    MaxValue = Excel.WorksheetFunction.Max(ws.Range("K2:K" & LastTicker))

    MinValue = Excel.WorksheetFunction.Min(ws.Range("K2:K" & LastTicker))
    
    MaxVolume = Excel.WorksheetFunction.Max(ws.Range("L2:L" & LastTicker))
    
    ws.Range("Q2:Q3").NumberFormat = "0.00%"

    For i = 2 To LastTicker
    
        If ws.Cells(i, "K").Value = MaxValue Then
        
            MaxTicker = ws.Cells(i, "I").Value
            
        ElseIf ws.Cells(i, "K").Value = MinValue Then
        
            MinTicker = ws.Cells(i, "I").Value
            
        End If
        
    Next i
    
        
    ws.Cells(2, "P").Value = MaxTicker
    ws.Cells(2, "Q").Value = MaxValue
    
    ws.Cells(3, "P").Value = MinTicker
    ws.Cells(3, "Q").Value = MinValue
    
    For i = 2 To LastTicker
        
        If ws.Cells(i, "L").Value = MaxVolume Then
        
            VolTicker = ws.Cells(i, "I").Value
            
        End If
        
    Next i
    
    ws.Cells(4, "P").Value = VolTicker
    ws.Cells(4, "Q").Value = MaxVolume
    
    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(4, "O").Value = "Greatest Total Volume"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
    
    ws.Columns("O:Q").AutoFit
    
    
    For k = 2 To LastTicker
    
        If ws.Cells(k, "J").Value < 0 Then
        
            ws.Range("J" & k).Interior.ColorIndex = 3
            
        ElseIf ws.Cells(k, "J").Value > 0 Then
        
            ws.Range("J" & k).Interior.ColorIndex = 4
            
        End If
        
            
    Next k
        
    

        
Next ws

    


End Sub
