Attribute VB_Name = "Module1"
' WILLIAM VANN HOMEWORK #2

Sub HW2FINALLY():
    
    Dim TickerColumn, DateColumn, OpenColumn, HighColumn, LowColumn, CloseColumn, VolumeColumn As Integer
    Dim FirstRow, LastRow As Long
    
    Dim i, j, k As Long
    Dim TickerCount As Integer
    Dim SummaryRow, SummaryLastRow As Integer
    
    Dim Ticker, PrevTicker, NextTicker As String
    Dim GreatestPrecentIncreaseTicker, GreastestPercentDecreaseTicker, GreatestTotalVolumeTicker As String
          
    Dim OpenPrice, ClosePrice, YearlyChange, YearlyPercentChange, GreatestPercentIncrease, GreatestPercentDecrease As Double
    
    Dim TotalVolume, GreatestTotalVolume As LongLong
    
    For Each ws In Worksheets
                
        ' Initialize all variables for each new worksheet
        
        FirstRow = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        TickerColumn = 1
        DateColumn = 2
        OpenColumn = 3
        HighColumn = 4
        LowColumn = 5
        CloseColumn = 6
        VolumeColumn = 7
        
        TickerCount = 0
        TotalVolume = 0
        YearlyChange = 0
        YearlyPercentChange = 0
        
        GreatestPercentIncrease = 0
        GreatestPercentDecrease = 0
        GreatestTotalVolume = 0
        
        GreatestIncreaseRowNum = 0
        GreatestDecreaseRowNum = 0
        GreatestVolRowNum = 0
        
        GreatestPercentIncreaseTicker = ""
        GreatestPercentDecreaseTicker = ""
        GreatestTotalVolumeTicker = ""
                      
        ' Start summary data at row 2
        
        SummaryRow = 2
        
            For i = FirstRow To LastRow
            
                Ticker = ws.Cells(i, TickerColumn).Value
                PrevTicker = ws.Cells(i - 1, TickerColumn).Value
                NextTicker = ws.Cells(i + 1, TickerColumn).Value
                
                If Ticker <> PrevTicker Then
                
                    ' First row of a ticker
                
                    TickerCount = TickerCount + 1
                    
                    TotalVolume = 0
                    YearlyChange = 0
                    YearlyPercentChange = 0
                            
                    OpenPrice = CDbl(ws.Cells(i, OpenColumn).Value)
                    TotalVolume = TotalVolume + CLngLng(ws.Cells(i, VolumeColumn).Value)
                                        
                ElseIf (Ticker = PrevTicker) And (Ticker <> NextTicker) Then
                
                    ' Last row of a ticker
                
                    ClosePrice = CDbl(ws.Cells(i, CloseColumn).Value)
                    TotalVolume = TotalVolume + CLngLng(ws.Cells(i, VolumeColumn).Value)
                    
                    YearlyChange = ClosePrice - OpenPrice
                    YearlyPercentChange = (YearlyChange / OpenPrice) * 100
                                                           
                    ' Write summary data to sheet
                    
                    ws.Cells(1, 9).Value = "Ticker"
                    ws.Cells(1, 10).Value = "Yearly Change"
                    ws.Cells(1, 11).Value = "Percent Change"
                    ws.Cells(1, 12).Value = "Total Stock Volume"
                    
                    ws.Cells(SummaryRow, 9).Value = Ticker
                    ws.Cells(SummaryRow, 10).Value = YearlyChange
                    ws.Cells(SummaryRow, 11).Value = YearlyPercentChange
                    ws.Cells(SummaryRow, 12).Value = TotalVolume
                    
                    ' Formatting
                                        
                    If YearlyChange > 0 Then
                    
                        ws.Cells(SummaryRow, 10).Interior.ColorIndex = 10
                        
                    ElseIf YearlyChange < 0 Then
                        
                        ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                                      
                    End If
                    
                    ' Increment SummaryRow
                    
                    SummaryRow = SummaryRow + 1
                   
                Else
                
                    ' Middle row of a ticker
                    
                    TotalVolume = TotalVolume + CLngLng(ws.Cells(i, VolumeColumn).Value)
                                    
                End If
                        
            Next i
            
            SummaryLastRow = SummaryRow
                                    
            ' Write Aggregate Summary Data
            
            ws.Range("J2:J" & SummaryLastRow).NumberFormat = "$0.00"
            ws.Range("K2:K" & SummaryLastRow).NumberFormat = "0.00%"
            ws.Range("L2:L" & SummaryLastRow).NumberFormat = "###,###,###,###,###,###"
            
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("Q4").NumberFormat = "###,###,###,###,###,###"
            
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            
            GreatestPercentIncrease = WorksheetFunction.Max(ws.Range("K2:K" & SummaryLastRow))
            GreatestPercentDecrease = WorksheetFunction.Min(ws.Range("K2:K" & SummaryLastRow))
            GreatestTotalVolume = WorksheetFunction.Max(ws.Range("L2:L" & SummaryLastRow))
           
            For j = 2 To SummaryLastRow
            
                If ws.Cells(j, 11).Value = GreatestPercentIncrease Then
                
                    ' write ticker and value
                        
                    ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
                    ws.Cells(2, 17).Value = GreatestPercentIncrease
                                  
                ElseIf ws.Cells(j, 11).Value = GreatestPercentDecrease Then
                
                    ' write ticker and value
                    
                    ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
                    ws.Cells(3, 17).Value = GreatestPercentDecrease
               
                ElseIf ws.Cells(j, 12).Value = GreatestTotalVolume Then
                
                    ' write ticker and value
                    
                    ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
                    ws.Cells(4, 17).Value = GreatestTotalVolume
               
                End If
              
            Next j
       
    ' Autofit all columns in sheet
        
    ws.Columns("A:Q").AutoFit
    
    Next ws
             
    Debug.Print ("The End!")
   
End Sub

