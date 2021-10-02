Attribute VB_Name = "Module1"
Sub StockMarket():

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets

        Dim Ticker As String
        Dim TotalStockVol As LongLong
        Dim LastRow As Long
        Dim SummaryTableRow As Integer
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        
        TotalStockVol = 0
        OpeningPrice = 0
        ClosingPrice = 0
        YearlyChange = 0
        PercentChange = 0
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        SummaryTableRow = 2
        OpeningPrice = Range("C2")
            
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
        
        For Row = 2 To LastRow
            If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
                Ticker = Range("A" & Row).Value
                ClosingPrice = Cells(Row, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice <> 0 Then
                    PercentChange = (YearlyChange / OpeningPrice)
                End If
                TotalStockVol = TotalStockVol + Range("G" & Row).Value
                Range("I" & SummaryTableRow).Value = Ticker
                Range("J" & SummaryTableRow).Value = YearlyChange
                If (YearlyChange > 0) Then
                    Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                ElseIf (YearlyChange <= 0) Then
                    Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
                Range("K" & SummaryTableRow).Value = FormatPercent(PercentChange, 2)
                Range("L" & SummaryTableRow).Value = TotalStockVol
                SummaryTableRow = SummaryTableRow + 1
                TotalStockVol = 0
                YearlyChange = 0
                PercentChange = 0
                OpeningPrice = Cells(Row + 1, 3).Value
            Else
                TotalStockVol = TotalStockVol + Range("G" & Row).Value
            End If
        Next Row
    
    Next ws
    
End Sub

