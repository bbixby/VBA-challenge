'Create a script that will loop through all the stocks for one year and output the following information
'The ticker symbol: Column A (1)
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
'--Open price Column C (3) from first appearance of ticker; Close Price Column F (6) from last appearance of ticker
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'--Calculation Close/Open - 1; format as percent
'The total stock volume of the stock.
'--Sum Column G (7)

Sub StockSummary()

' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------

For Each ws In Worksheets

' --------------------------------------------
'SET VARIABLES
' --------------------------------------------
'Ticker Name
Dim TickerSym As String
'New Open Price to set new row open price; initialize to first open price
Dim NewOpenPrice As Double
NewOpenPrice = ws.Cells(2, 3).Value
'Ticker Open Price Column C (3)
Dim OpenPrice As Double
'Ticker Close Price Column F (6)
Dim ClosePrice As Double
'Yearly change in price from beginning open to end close; ClosePrice - OpenPrice
Dim YearlyChange As Double
'Percent change; ClosePrice divided by Open Price - 1
Dim PercentChange As Double
'Volume Total sum Column G (7); initialize with 0
Dim VolumeTotal As LongLong
VolumeTotal = 0
'Summary table last row for pastes; start at 2 to make room for header
Dim SummaryTableRow As Integer
SummaryTableRow = 2
'Retain the Greatest INCREASE symbol and percent; initialize with 0 for ticker comparisons NOTE assumes change > 0
Dim IncreaseTickerSym As String
Dim IncreasePercentChange As Double
IncreasePercentChange = 0
'Retain the Greatest DECREASE symbol and percent; initialize with 0 for ticker comparisons NOTE assumes change < 0
Dim DecreaseTickerSym As String
Dim DecreasePercentChange As Double
DecreasePercentChange = 0
'Retain the Greatest VOLUME symbol and percent; initialize with 0 for ticker comparisons NOTE assumes volume > 0
Dim VolumeTickerSym As String
Dim VolumeGreatest As LongLong
VolumeGreatest = 0

' --------------------------------------------
'PRINT HEADERS
' --------------------------------------------
'Print the Summary Table Headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
'Print the Greatest table headers and row labels
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

' --------------------------------------------
'LOOP ENTIRE DATA SET
' --------------------------------------------
'Find the last row for the iteration
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Look through all stock prices; start with 2 to skip header
For i = 2 To LastRow

    'If starting a new ticker...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Add the last volume
        VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
        'Set the New Open Price to the row below (i + 1)
        NewOpenPrice = ws.Cells(i + 1, 3).Value
        'Set the ticker name
        TickerSym = ws.Cells(i, 1).Value
        'Set the Close Price
        ClosePrice = ws.Cells(i, 6).Value
        'Calculate the YearlyChange
        YearlyChange = ClosePrice - OpenPrice
        'Calculate the PriceChange; check against 0 values in numerator and denominator
        If ClosePrice = 0 And OpenPrice > 0 Then PercentChange = -1
        If ClosePrice > 0 And OpenPrice = 0 Then PercentChange = 1
        If ClosePrice = 0 And OpenPrice = 0 Then PercentChange = 0
        If ClosePrice <> 0 And OpenPrice <> 0 Then PercentChange = ClosePrice / OpenPrice - 1
        
            'Print values
            ws.Range("I" & SummaryTableRow).Value = TickerSym
            ws.Range("J" & SummaryTableRow).Value = YearlyChange
                'set the YearlyChange cell green if positive, red if negative
                If YearlyChange > 0 Then ws.Range("J" & SummaryTableRow).Interior.Color = vbGreen
                If YearlyChange < 0 Then ws.Range("J" & SummaryTableRow).Interior.Color = vbRed
            ws.Range("K" & SummaryTableRow).Value = PercentChange
                'Check Percent Change against Greatest values; if surpasses, retain value and ticker
                'INCREASE Check
                If PercentChange > IncreasePercentChange Then
                    IncreaseTickerSym = TickerSym
                    IncreasePercentChange = PercentChange
                'DECREASE Check
                ElseIf PercentChange < DecreasePercentChange Then
                    DecreaseTickerSym = TickerSym
                    DecreasePercentChange = PercentChange
                End If
            'Format Percent Change value as percentage
            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
            ws.Range("L" & SummaryTableRow).Value = VolumeTotal
                'Check against Greatest Volume; if surpasses, retain value and ticker
                'VOLUME Check
                If VolumeTotal > VolumeGreatest Then
                    VolumeTickerSym = TickerSym
                    VolumeGreatest = VolumeTotal
                End If
            'Format all printed columns to auto fit width
            ws.Range("I:L").Columns.AutoFit
            
        'Add 1 to the SummaryTableRow
        SummaryTableRow = SummaryTableRow + 1
        'Reset the VolumeTotal
        VolumeTotal = 0
    
    'if seeing the same ticker...
    ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
    
        'Set the New Open Price to Open Price
        OpenPrice = NewOpenPrice
        'Add volume
        VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
        
    End If
    
Next i

' --------------------------------------------
'PRINT GREATEST VALUES
' --------------------------------------------
'Print the Greatest tickers and values
ws.Cells(2, 15).Value = IncreaseTickerSym
ws.Cells(2, 16).Value = IncreasePercentChange
ws.Cells(3, 15).Value = DecreaseTickerSym
ws.Cells(3, 16).Value = DecreasePercentChange
ws.Cells(4, 15).Value = VolumeTickerSym
ws.Cells(4, 16).Value = VolumeGreatest

'Format Greatest values
ws.Range("P2:P3").NumberFormat = "0.00%"
ws.Cells(4, 16).NumberFormat = "General"
ws.Range("N:P").Columns.AutoFit

Next ws

End Sub



