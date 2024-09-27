Attribute VB_Name = "Module1"
Sub StockProgram()

'Visual Basic Code - Module #2
'Jose Moncada

'General Variables
Dim OpenValue As Double
Dim CloseValue As Double
Dim Delta As Double
Dim PercentChange As Double
Dim StockVolume As Double
Dim MaxPercentChange As Double
Dim MinPercentChange As Double
Dim MaxStockVolume As Double
Dim StockNameMaxPercent As String
Dim StockNameMinPercent As String
Dim StockNameMaxVolume As String

'Added variable for Each WorkSheet
Dim WS As Worksheet

For Each WS In ThisWorkbook.Sheets
WS.Activate

'Variable to Find Last Row
Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'StartRow starts at 2 because thats where the first stock starts
StartRow = 2

'The Row where my summary table will start.
SummaryRow = 2

'Variable to hold Stockvolume Start Value
StockVolume = 0

'Looop through rows from 2 to 100000
'Check if Stock changes or it is the last row
For i = 2 To LastRow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Or i = LastRow Then

'StockName will keep the StartRow value
Stock = Cells(StartRow, 1).Value

'Add to Stockvolume
StockVolume = StockVolume + Cells(i, 7).Value

'Open Value from first Row
OpenValue = Cells(StartRow, 3).Value

'CloseValue from row before stock changes
CloseValue = Cells(i, 6).Value

'Delta Open VS Close
Delta = CloseValue - OpenValue

'Percent Change
PercentChange = (CloseValue / OpenValue) - 1

'Print Column Titles
Cells(1, 9).Value = "Title"
Cells(1, 10).Value = "Quarterly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'Print Results
Cells(SummaryRow, 9).Value = Stock
Cells(SummaryRow, 10).Value = Delta
Cells(SummaryRow, 11).Value = PercentChange
Cells(SummaryRow, 12).Value = StockVolume

'Check for Max/Min Percent Change and Max Stock Volume
        If SummaryRow = 2 Or PercentChange > MaxPercentChange Then
            MaxPercentChange = PercentChange
            StockNameMaxPercent = Stock
        End If

        If SummaryRow = 2 Or PercentChange < MinPercentChange Then
            MinPercentChange = PercentChange
            StockNameMinPercent = Stock
        End If

        If SummaryRow = 2 Or StockVolume > MaxStockVolume Then
            MaxStockVolume = StockVolume
            StockNameMaxVolume = Stock
        End If

'Add conditional format to quarterly Change column.

With Cells(SummaryRow, 10) 'Column J is the 10th column
    If .Value > 0 Then
        .Interior.Color = RGB(0, 255, 0) 'Green for positive values
    ElseIf .Value < 0 Then
        .Interior.Color = RGB(255, 0, 0) 'Red for negative values
    End If
End With

With Cells(SummaryRow, 11) 'Column K is the 11th column
    If .Value > 0 Then
        .Interior.Color = RGB(0, 255, 0) 'Green for positive values
    ElseIf .Value < 0 Then
        .Interior.Color = RGB(255, 0, 0) 'Red for negative values
    End If
End With

'Next line in summary row
SummaryRow = SummaryRow + 1

'Next Start row
StartRow = i + 1
'Reset StockVolume
StockVolume = 0

Else
StockVolume = StockVolume + Cells(i, 7).Value

End If
Next i

'Print Max/Min Percent Change and Max Stock Volume
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(2, 17).Value = MaxPercentChange
Cells(3, 17).Value = MinPercentChange
Cells(4, 17).Value = MaxStockVolume

'Print Stock names associated with those values
Cells(2, 16).Value = StockNameMaxPercent
Cells(3, 16).Value = StockNameMinPercent
Cells(4, 16).Value = StockNameMaxVolume

Next WS
End Sub

