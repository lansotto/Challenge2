Attribute VB_Name = "Module1"
Sub Challenge()
    
For Each xsheet In ThisWorkbook.Worksheets
        xsheet.Select
    
        Dim NewRowCount As Long
        NewRowCount = "1"
        Dim firstday As Long
        Dim lastday As Long
        firstday = Left(Range("B2"), 4) & "0102"
        lastday = Left(Range("B2"), 4) & "1231"
    
        Dim dateInColumnB As Long
        Dim openingvalue As Double
        Dim closingvalue As Double
        Dim ticker As String
    
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("K:K").NumberFormat = "0.00%"
        Range("J:J").NumberFormat = "0.00"
        Range("Q2:Q3").NumberFormat = "0.00%"
        
        Dim LastrowA As Long
        Dim LastrowB As Long
        LastrowA = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
        LastrowB = ActiveSheet.Cells(Rows.Count, 10).End(xlUp).Row
    
       For j = 2 To LastrowA

       dateInColumnB = Cells(j, 2).Value
           ticker = Cells(j, 1).Value
           'Dim result As Long
           'result = Application.WorksheetFunction.CountIf(Range("I2:I" & j), ticker)
           'If result = 0 Then
           'I was going to use the following: but that assumes that the tickers are sorted alphabetically
           If Cells(j, 1).Value <> Cells(j - 1, 1).Value Then
           NewRowCount = NewRowCount + 1
           
           Range("I" & NewRowCount).Value = ticker
           
                If dateInColumnB = firstday Then
                openingvalue = Cells(j, 3).Value
                Else
                End If
           
            ElseIf dateInColumnB = lastday Then

            closingvalue = Cells(j, 6).Value

            Range("J" & NewRowCount).Value = closingvalue - openingvalue
                If Range("J" & NewRowCount).Value < 0 Then
                Range("J" & NewRowCount).Interior.ColorIndex = 3
                ElseIf Range("J" & NewRowCount).Value > 0 Then
                Range("J" & NewRowCount).Interior.ColorIndex = 4
                End If
            
            Range("K" & NewRowCount).Value = (closingvalue - openingvalue) / openingvalue
            
                If Range("K" & NewRowCount).Value < 0 Then
                Range("K" & NewRowCount).Interior.ColorIndex = 3
                ElseIf Range("K" & NewRowCount).Value > 0 Then
                Range("K" & NewRowCount).Interior.ColorIndex = 4
                End If
            Range("L" & NewRowCount).Value = Application.WorksheetFunction.SumIf(Range("A2:A" & LastrowA), ticker, Range("G2:G" & LastrowA))
            End If
           
           NewRowCount = NewRowCount
        Next j
        
        Range("Q2").Value = Application.WorksheetFunction.Max(Range("k:k"))
        Range("Q3").Value = Application.WorksheetFunction.Min(Range("k:k"))
        Range("Q4").Value = Application.WorksheetFunction.Max(Range("l:l"))

        Dim GreatestIncreaseRow As Long
        GreatestIncreaseRow = WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("k:k")), Range("K:K"), 0)
        Range("P2").Value = Range("I" & GreatestIncreaseRow).Value
        
        Dim GreatestDecreaseRow As Long
        GreatestDecreaseRow = WorksheetFunction.Match(Application.WorksheetFunction.Min(Range("k:k")), Range("K:K"), 0)
        Range("P3").Value = Range("I" & GreatestDecreaseRow).Value
        
        Dim GreatestVolumeRow As Long
        GreatestVolumeRow = WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("L:L")), Range("L:L"), 0)
        Range("P4").Value = Range("I" & GreatestVolumeRow).Value

        Next xsheet

      MsgBox ("Done")
End Sub


