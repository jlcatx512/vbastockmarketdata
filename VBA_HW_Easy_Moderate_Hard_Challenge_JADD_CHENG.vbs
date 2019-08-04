Sub vba_hw_easy_moderate_hard_challenge()
' 1. DECLARE WORKSHEET OBJECT VARIABLE FOR FOR-EACH LOOP
    Dim ws As Worksheet
    'Start For-Each loop to go through each workbook worksheet.
    For Each ws In Worksheets
        ' START CODE TO EXECUTE ON EACH WORKSHEET.
' 2. DECLARE VARIABLES AND ASSIGN INITIAL VALUES.
        Dim tickerSymbol As String
        tickerSymbol = ws.Range("A2").Value
        Dim totalVolume As Double
        totalVolume = ws.Range("G2").Value
        Dim begYearOpenPrice As Double
        begYearOpenPrice = ws.Range("C2").Value
        Dim endYearClosePrice As Double
        endYearClosePrice = ws.Range("F2").Value
        Dim summaryTableRow As Integer
        summaryTableRow = 2 ' Initialize at second row because of header.
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim lRow As Long ' Last row of original data set.
        lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
' 3. SET UP SUMMARY TABLE HEADERS.
    ' Summary Table 1
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    ' Summary Table 2
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
' 4. LOOP THROUGH EVERY ROW OF TWO DATA SETS
    ' Calculate total stock volume and yearly change grouped by ticker symbol.
        For I = 2 To lRow
            'The if statement tests for the last instance of a ticker symbol.
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                ' Set Ticker Symbol.
                tickerSymbol = ws.Cells(I, 1).Value
                ' Print ticker symbol to summary table.
                ws.Range("I" & summaryTableRow).Value = tickerSymbol
                ' Set value of end of year closing price.
                endYearClosePrice = ws.Cells(I, 6).Value
                ' Set value of Total Stock Volume.
                totalVolume = totalVolume + ws.Cells(I, 7).Value
                ' Print total stock volume to summary table.
                ws.Range("L" & summaryTableRow).Value = totalVolume
                ' Calculate yearly change.
                yearlyChange = endYearClosePrice - begYearOpenPrice
                ' Print yearly change to summary table.
                ws.Range("J" & summaryTableRow).Value = yearlyChange
                ' Calculate percent change. Nested if statement to catch beginning year opening prices of 0.
                If begYearOpenPrice = 0 Then
                    ' percentChange = 0 ' I set this to zero and not N/A to avoid an error calculating percent change. A better way might be to use a error handling statement.
                    percentChange = 0
                Else
                    percentChange = yearlyChange / begYearOpenPrice
                End If
                ' Print percent change to summary table.
                ws.Range("K" & summaryTableRow).Value = percentChange
                ' Format percent change as %.
                ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
                ' Format yearly change cells conditionally. Green for positive. Red for negative.
                If yearlyChange > 0 Then
                    ws.Range("J" & summaryTableRow).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlyChange < 0 Then
                    ws.Range("J" & summaryTableRow).Interior.Color = RGB(255, 0, 0)
                Else ' Zero
                End If
                ' Reset totalVolume for next ticker symbol.
                totalVolume = 0
                ' Increment +1 summary table row.
                summaryTableRow = summaryTableRow + 1
            ' Elseif tests for the first instance of a ticker symbol    
            Elseif ws.Cells(I - 1, 1).Value <> ws.Cells(I, 1).Value Then
                ' If first cell of ticker symbol, assign value to begYearOpenPrice.
                begYearOpenPrice = ws.Cells(I, 3).Value
                ' Add value in volume column to totalVolume.
                totalVolume = totalVolume + ws.Cells(I, 7).Value
            ' Else tests for all other instances of a ticker symbol, i.e. neither first nor last.
            Else
                ' Add value in volume column to totalVolume.
                totalVolume = totalVolume + ws.Cells(I, 7).Value
            End If
        Next I
' 5. CALCULATE GREATEST TOTAL STOCK VOLUME, % INCREASE/DECREASE
        ' HARD SECTION VARIABLES
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim lRow2 As Long
        lRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row ' last row of first summary table.
        Dim tickerGreatestIncrease As String
        Dim tickerGreatestDecrease As String
        Dim tickerGreatestVolume As String
        greatestDecrease = ws.Range("K2").Value
        greatestIncrease = ws.Range("K2").Value
        greatestVolume = ws.Range("L2").Value
        tickerGreatestIncrease = ws.Range("I2").Value
        tickerGreatestDecrease = ws.Range("I2").Value
        tickerGreatestVolume = ws.Range("L2").Value
        ' Loop through summary table 1.
        For j = 2 To lRow2
            If ws.Range("K" & j + 1).Value < greatestDecrease Then
                greatestDecrease = ws.Range("K" & j + 1).Value
                tickerGreatestDecrease = ws.Range("I" & j + 1).Value
            ElseIf ws.Range("K" & j + 1).Value > greatestIncrease Then
                greatestIncrease = ws.Range("K" & j + 1).Value
                tickerGreatestIncrease = ws.Range("I" & j + 1).Value
            ElseIf ws.Range("L" & j + 1).Value > greatestVolume Then
                greatestVolume = ws.Range("L" & j + 1).Value
                tickerGreatestVolume = ws.Range("I" & j + 1).Value
            Else
            End If
        Next j
        ' Print results to summary table 2.
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("Q3").Value = greatestDecrease
        ws.Range("Q4").Value = greatestVolume
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("P2").Value = tickerGreatestIncrease
        ws.Range("P3").Value = tickerGreatestDecrease
        ws.Range("P4").Value = tickerGreatestVolume
        ' Auto-fit columns to best fit.
        ws.Range("I:Q").Columns.AutoFit
    ' Iterate to next worksheet.
    Next ws
End Sub