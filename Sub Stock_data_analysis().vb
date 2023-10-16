Sub Stock_data_analysis()

    Dim TickerS As String
    Dim LastRow As Long
    Dim i As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalVolume As LongLong
    Dim Datarow As Long
    Dim ws As Worksheet
    
    'Dim max_INC  As Long
   ' Dim tag_INC As Long
    'Dim min_DEA As Long
    'Dim tag_DEA As Long
   ' Dim max_Total As Long
   ' Dim tag_Total As Long
    
    For Each ws In Worksheets

    
    'name first rows & table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Value"
        ws.Range("O1").Value = "Description"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Value"
        
    ' start Data row
        Datarow = 2
    
    'Find last row
                 
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Identify the first open price
        OpenPrice = ws.Cells(2, 3).Value
        
            
    ' For finding unique ticker symbol

        For i = 2 To LastRow
        
            TickerS = ws.Cells(i, 1).Value
            ClosePrice = ws.Cells(i, 6).Value
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                 
            If ws.Cells(i + 1, 1).Value <> TickerS Then
                YearlyChange = ClosePrice - OpenPrice
    'Calculate percentchange
                If OpenPrice <> 0 Then
                    PercentageChange = (YearlyChange / OpenPrice) * 100
                Else
                    PercentageChange = 0
                End If
                
    'Place values in new rows
            ws.Cells(Datarow, 9).Value = TickerS
            ws.Cells(Datarow, 10).Value = YearlyChange
            ws.Cells(Datarow, 11).Value = PercentageChange
            ws.Range("K" & Datarow).NumberFormat = "0.00%"
            ws.Cells(Datarow, 12).Value = TotalVolume
            OpenPrice = ws.Cells(i + 1, 3).Value
            TotalVolume = 0
    'Formating cell fill colour for Percentage and Yearly Change
                If (YearlyChange < 0) Then
                    ws.Range("J" & Datarow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Datarow).Interior.ColorIndex = 4
                End If
                
                If (PercentageChange < 0) Then
                    ws.Range("K" & Datarow).Interior.ColorIndex = 3
                Else
                    ws.Range("K" & Datarow).Interior.ColorIndex = 4
                End If
                Datarow = Datarow + 1
                
                
    End If
                                    
        Next i
            ' take the max and min
                ws.Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & LastRow)) * 100
                ws.Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & LastRow)) * 100
                ws.Range("Q4") = WorksheetFunction.Max(Range("L2:L" & LastRow))

    ' returns start from second row
                increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
                decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
                volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & LastRow)), Range("L2:L" & LastRow), 0)

    ' final ticker symbol for  total, greatest % of increase and decrease, and average
                ws.Range("P2") = Cells(increase_number + 1, 9)
                ws.Range("P3") = Cells(decrease_number + 1, 9)
                ws.Range("P4") = Cells(volume_number + 1, 9)
                
    Next ws
               
        
End Sub
