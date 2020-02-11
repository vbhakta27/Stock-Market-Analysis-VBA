Attribute VB_Name = "Module2"
'Loops through all worksheets
Sub WorksheetLoop()

Dim ws As Worksheet

    For Each ws In Worksheets
        
        'Activates worksheet
        ws.Activate
        'Runs Stock Analysis on active worksheet and autofits columns
        Call StockAnalysis
        ws.Columns("I:P").AutoFit
    Next

End Sub


'Main program that analyzes every stock
Sub StockAnalysis()

'Create headers for table
Cells(1, 9).Value = "Tickers"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Define variables
Dim Ticker As String
Dim LastRowData As Long
Dim Volume As Double
Dim I As Long
Dim OpenDate As Long
Dim CloseDate As Long
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double

'Get number of last row that contains data
LastRowData = Range("A" & Rows.Count).End(xlUp).Row

'Start table row at 2
SummaryTableRow = 2

'Set inital Volume to zero
Volume = 0

'Set Initial OpenPrice for first ticker
OpenPrice = Cells(2, 3).Value

'Loops through data set to last row
    For I = 2 To LastRowData
    
        'Checks if next ticker symbol is equal to the ticker symbol we are on
        If Cells(I, 1) <> Cells(I + 1, 1) Then
            
            'Sets Ticker as previous symbol
            Ticker = Cells(I, 1)
            
            'Adds on to Volume
            Volume = Volume + Cells(I, 7)
            
            'Grabs ClosePrice
            ClosePrice = Cells(I, 6).Value
            
            'Calculates YearlyChange
            YearlyChange = ClosePrice - OpenPrice
            
                'Calculates PercentChange
                If ClosePrice = 0 Or OpenPrice = 0 Then
                    YearlyChange = 0
                
                Else
                    PercentChange = (YearlyChange / OpenPrice)
                
                End If
                
            'Prints Ticker into table
            Cells(SummaryTableRow, 9) = Ticker
            
            'Prints YearlyChange
            Cells(SummaryTableRow, 10) = YearlyChange
            
                'Check to see if YearlyChange is positive then make the cell green or else make it red
                If Cells(SummaryTableRow, 10).Value >= 0 Then
                    Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                
                Else
                    Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
                
                End If
            
            'Prints PercentChange as a percent into table
            formattedPercentChange = Format(PercentChange, "0.00%")
            Cells(SummaryTableRow, 11) = formattedPercentChange
            
            'Prints Volume into table
            Cells(SummaryTableRow, 12) = Volume
            
            'Adds one  to summary table row so it will go to the next row
            SummaryTableRow = SummaryTableRow + 1
            
            'Set OpenPrice to for next ticker symbol
            OpenPrice = Cells(I + 1, 3).Value
            
            'Reset Volume
            Volume = 0
            
        Else
            'Adds on to Volume
            Volume = Volume + Cells(I, 7).Value
        
        End If
        
    Next I
    


'Create column and row headers for new table
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"


'Sets all inital variables to zero
GreatestPercentIncrease = 0
GreatestPercentDecrease = 0
GreatestTotalVolume = 0

'Reset SummaryTableRow to last table row
SummaryTableRow = SummaryTableRow - 1

    'Loop through new table that was made in earlier loop
    For I = 2 To SummaryTableRow

        'Checks to see if current row % change is greater than last saved GreatestPercentIncrease
        If Cells(I, 11).Value > GreatestPercentIncrease Then
            GreatestPercentIncrease = Cells(I, 11).Value
            GreatestPrecentIncreaseTicker = Cells(I, 9)
        End If
    
        'Checks to see if current row % change is less than last saved GreatestPercentDecrease
        If Cells(I, 11).Value < GreatestPercentDecrease Then
            GreatestPercentDecrease = Cells(I, 11).Value
            GreatestPrecentDecreaseTicker = Cells(I, 9)
        End If

        'Checks to see if current row volume is greater than last saved GreatestTotalVolume
        If Cells(I, 12).Value > GreatestTotalVolume Then
            GreatestTotalVolume = Cells(I, 12).Value
            GreatestVolumeTicker = Cells(I, 9)
        End If

    Next I

'Prints all values on table
Cells(2, 16).Value = GreatestPrecentIncreaseTicker
Cells(3, 16).Value = GreatestPrecentDecreaseTicker
Cells(4, 16).Value = GreatestVolumeTicker
'Converts values into percentages
formattedGreatestPercentIncrease = Format(GreatestPercentIncrease, "0.00%")
formattedGreatestPercentDecrease = Format(GreatestPercentDecrease, "0.00%")
Cells(2, 17).Value = formattedGreatestPercentIncrease
Cells(3, 17).Value = formattedGreatestPercentDecrease
Cells(4, 17).Value = GreatestTotalVolume



End Sub
