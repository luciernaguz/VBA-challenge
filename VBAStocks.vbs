
Sub VBAStock()
'VBA scripting to analyze real stock market data
Dim WS As Worksheet
'WS=worksheet

    'loop all sheets all file
    For Each WS In Worksheets
    
        Dim AssigR As Long
        Dim NumberR As Long
        'Declare variables
        Dim TickerInitial As String
        TickerInitial = ""
        Dim BegginingNumber As Double
        BegginingNumber = 0
        Dim EndingNumber As Double
        EndingNumber = 0
        Dim YearlyChange As Double
        YearlyChange = 0
        Dim PercentChange As Double
        PercentChange = 0
        Dim TStockVolume As Double
        TStockVolume = 0
        Dim GreatesIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestTotal As Double
        GreatestTotal = 0
        Dim TickerIncrease As String
        Dim TickerDecrease As String
        Dim TotalTicker As String
        TickerIncrease = ""
        TickerDecrease = ""
        TotalTicker = ""
        
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
        WS.Cells(1, 15).Value = "Ticker"
        WS.Cells(1, 16).Value = "Value"
        WS.Cells(2, 14).Value = "Greatest % Increase"
        WS.Cells(3, 14).Value = "Greatest % Decrease"
        WS.Cells(4, 14).Value = "Greatest Total Volume"
        
        
        'Counts the number of rows
        NumberR = WS.Cells(Rows.Count, 1).End(xlUp).Row
        AssigR = 2
        BegginingNumber = WS.Cells(2, 3).Value
        For i = 2 To NumberR
                'Iteration finding values
                If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
            
                    'Ticker Initial
                    TickerInitial = WS.Cells(i, 1).Value
                    
                    'Yearly Change
                    EndingNumber = WS.Cells(i, 6).Value
                    YearlyChange = (EndingNumber - BegginingNumber)
                               
                    'Percent Change
                    If BegginingNumber <> 0 Then
                    PercentChange = (YearlyChange / BegginingNumber) * 100
                    End If
                                                         
                    'Total Stock Volume
                    TStockVolume = TStockVolume + WS.Cells(i, 7).Value
                    
                    'Put Initials found in Column9
                    WS.Cells(AssigR, 9).Value = TickerInitial
                    ' Put the yearly change volume stock
                    WS.Cells(AssigR, 10).Value = YearlyChange
                    ' Put Percent Change
                    WS.Cells(AssigR, 11).Value = PercentChange
                    WS.Cells(AssigR, 11).Value = Application.WorksheetFunction.Round(PercentChange, 2) & "%"
                    ' Put and sum the total volume stock
                    WS.Cells(AssigR, 12).Value = TStockVolume
                                   
                                   
                    'Assigning color according to the value YearChange
                    If (YearlyChange > 0) Then
                        WS.Cells(AssigR, 10).Interior.ColorIndex = 4
                    ElseIf (YearlyChange <= 0) Then
                       WS.Cells(AssigR, 10).Interior.ColorIndex = 3
                    End If
                    
                    'Greatest Increase/Decrease
                    If (PercentChange > GreatestIncrease) Then
                        GreatestIncrease = PercentChange
                        TickerIncrease = TickerInitial
                     
                    ElseIf (PercentChange < GreatestDecrease) Then
                        GreatestDecrease = PercentChange
                        TickerDecrease = TickerInitial
                    End If
                    
                    'Greatest Volume
                    If (TStockVolume > GreatestTotal) Then
                        GreatestTotal = TStockVolume
                        TotalTicker = TickerInitial
                    End If
                    
                    
                    'values to continue with next ticker
                    AssigR = AssigR + 1
                    BegginingNumber = WS.Cells(i + 1, 3).Value
                    
                    ' Clear values YearlyChange = 0
                    PercentChange = 0
                    TStockVolume = 0
                    
                    Else
                    TStockVolume = TStockVolume + WS.Cells(i, 7).Value
                
                End If
                
                
        Next i
                   'printing final values
                   WS.Cells(2, 15).Value = TickerIncrease
                   WS.Cells(2, 16).Value = GreatestIncrease
                   WS.Cells(2, 16).Value = Application.WorksheetFunction.Round(GreatestIncrease, 2) & "%"
                   'WS.Cells(2, 16).NumberFormat = "%"
                   WS.Cells(3, 15).Value = TickerDecrease
                   WS.Cells(3, 16).Value = GreatestDecrease
                   WS.Cells(3, 16).Value = Application.WorksheetFunction.Round(GreatestDecrease, 2) & "%"
                   'WS.Cells(3, 16).NumberFormat = "%"
                   WS.Cells(4, 15).Value = TotalTicker
                   WS.Cells(4, 16).Value = GreatestTotal
                   
    Next WS ' end WS loop
End Sub
