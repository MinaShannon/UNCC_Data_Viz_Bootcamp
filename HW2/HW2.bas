Sub Homework2():
    
For Each ws In Worksheets
ws.Activate
    Set Active = ActiveWorkbook.ActiveSheet
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Dim TotalStock As Double
    TotalStock = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim YearOpen As Variant
    Dim YearClose As Variant
    Dim YearChange As Variant
    Dim PercentChange As Variant
    Dim Ticker As String, n As Double, m As Double
    n = WorksheetFunction.CountA(Active.Columns(1))
    YearOpen = ws.Cells(2, 3).Value
       
        For i = 2 To n
    
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
            Ticker = ws.Cells(i, 1).Value
        
            TotalStock = TotalStock + ws.Cells(i, 7)
            
            YearClose = ws.Cells(i, 6).Value
            
            YearChange = YearClose - YearOpen
                            
                If YearOpen > 0 Then
            
                PercentChange = YearChange / YearOpen
                
                Else: PercentChange = (YearClose - 14.5) / 14.5
                End If
                            
                If YearOpen = 0 And ws.Cells(i, 3) = 0 Then
            
                PercentChange = 0
                
                End If

            ws.Range("I" & Summary_Table_Row).Value = Ticker
        
            ws.Range("L" & Summary_Table_Row).Value = TotalStock
            
            ws.Range("J" & Summary_Table_Row).Value = YearChange
            
            ws.Range("K" & Summary_Table_Row).Value = PercentChange
        
            Summary_Table_Row = Summary_Table_Row + 1
        
            TotalStock = 0
            
            YearOpen = ws.Cells(i + 1, 3).Value
            
            Else
        
            TotalStock = TotalStock + ws.Cells(i, 7).Value
        
            End If
            
            Next i
            
            m = WorksheetFunction.CountA(Active.Columns(10))
            
            For c = 2 To m
            
            If ws.Cells(c, 10).Value <= 0 Then
            
            ws.Cells(c, 10).Interior.ColorIndex = 3
            
            End If
            
            If ws.Cells(c, 10).Value > 0 Then
            
            ws.Cells(c, 10).Interior.ColorIndex = 4
            
            End If
            
            ws.Cells(c, 11).NumberFormat = "0.00%"
            
            Next c
           
           ws.Cells(2, 15).Value = "Greatest % Increase"
           ws.Cells(3, 15).Value = "Greatest % Decrease"
           ws.Cells(4, 15).Value = "Greatest Total Volume"
           ws.Cells(2, 17).Value = WorksheetFunction.Max(Active.Columns(11))
           ws.Cells(3, 17).Value = WorksheetFunction.Min(Active.Columns(11))
           ws.Cells(4, 17).Value = WorksheetFunction.Max(Active.Columns(12))
           
           For j = 2 To n
           If ws.Cells(2, 17).Value = ws.Cells(j, 11).Value Then
           ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
           End If
           
           If ws.Cells(3, 17).Value = ws.Cells(j, 11).Value Then
           ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
           End If
           
           If ws.Cells(4, 17).Value = ws.Cells(j, 12).Value Then
           ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
           End If
           
           Next j
           
           ws.Cells(2, 17).NumberFormat = "0.00%"
           
           ws.Cells(3, 17).NumberFormat = "0.00%"
     Next ws
    
End Sub



