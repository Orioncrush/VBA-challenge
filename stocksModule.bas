Attribute VB_Name = "Module1"
' Instructions
' Create a script that llops through all the stocks for one year and outputs the following information
'   - The ticker symbol
'   - Yearly change from the opening price to the closing price
'       - Conditional formatting for positive or negative change
'   - Percentage of change from the opening price
'   - Total Stock volume for the year

'   Bonus!
'       Add functionality that returns
'           - Greatest % increase
'           - Greatest % decrease
'           - Greatest Total Volume
Sub StockMarket()
    ' These variables are to track across a single stock
    Dim tickerSymbol As String
    Dim openPrice As Double
    Dim earliestDate As Double
    Dim closingPrice As Double
    Dim latestDate As Double
    Dim totalVolume As Double
    
    ' These variables are globals to track the leaders
    Dim greatestIncrease As Double
    Dim gITicker As String
    Dim greatestDecrease As Double
    Dim gDTicker As String
    Dim greatestVolume As Double
    Dim gVTicker As String
    
    ' This variable tracks which row to print on
    printRow = 2
    
    ' get the number of rows in the sheet
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Create headers for the colated and award data
    Call PrintHeaders
    
    ' Loop through each line of the book to gather information
    For i = 2 To lastRow
        ' Check to see if this is a new ticker
        If tickerSymbol = "" Then
            ' This is the first stock
            earliestDate = Cells(i, 2).Value
            latestDate = Cells(i, 2).Value
            tickerSymbol = Cells(i, 1).Value
            ' save opening and closing prices
            openPrice = Cells(i, 3).Value
            closingPrice = Cells(i, 6).Value
            ' add to volume
            totalVolume = Cells(i, 7).Value
        ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            ' WRITE PREVIOUS INFO AND RESET VARIABLES
            Call WriteStockRow(Int(printRow), tickerSymbol, openPrice, closingPrice, totalVolume)
            deltaPercent = Round(((closingPrice - openPrice) / openPrice) * 100, 2)
            If deltaPercent > greatestIncrease Then
                ' new award winner
                greatestIncrease = deltaPercent
                gITicker = tickerSymbol
            End If
            If deltaPercent < greatestDecrease Then
                ' new award winner
                greatestDecrease = deltaPercent
                gDTicker = tickerSymbol
            End If
            If totalVolume > greatestVolume Then
                ' new award winner
                greatestVolume = totalVolume
                gVTicker = tickerSymbol
            End If
            printRow = printRow + 1
            ' Save ticker symbol
            earliestDate = Cells(i, 2).Value
            latestDate = Cells(i, 2).Value
            tickerSymbol = Cells(i, 1).Value
            ' save opening and closing prices
            openPrice = Cells(i, 3).Value
            closingPrice = Cells(i, 6).Value
            ' add to volume
            totalVolume = Cells(i, 7).Value
        ' Otherwise
        Else
            ' Check if this information is a newer or older stock info
            If Cells(i, 2).Value > latestDate Then
                'Newer information
                latestDate = Cells(i, 2).Value
                closingPrice = Cells(i, 6).Value
                totalVolume = totalVolume + Cells(i, 7).Value
            ElseIf Cells(i, 2).Value < earliestDate Then
                'Older information
                earliestDate = Cells(i, 2).Value
                openPrice = Cells(i, 3).Value
                totalVolume = totalVolume + Cells(i, 7).Value
            Else
                totalVolume = totalVolume + Cells(i, 7).Value
                'Middle Information
            End If
        End If
        Next i
    deltaValue = closingPrice - openPrice
    Call WriteStockRow(Int(printRow), tickerSymbol, openPrice, closingPrice, totalVolume)
    
    'Greatest Increase
    Cells(2, 17).Value = greatestIncrease
    Cells(2, 16).Value = gITicker
    'Greatest Decrease
    Cells(3, 17).Value = greatestDecrease
    Cells(3, 16).Value = gDTicker
    'Greatest Volume
    Cells(4, 17).Value = greatestVolume
    Cells(4, 16).Value = gVTicker
End Sub

' This subroutine takes in information to write a single stock
Sub WriteStockRow(printRow As Integer, ticker As String, openPrice As Double, closePrice As Double, totalVolume As Double)
    Cells(printRow, 9).Value = ticker
    yearChange = closePrice - openPrice
    Cells(printRow, 10).Value = yearChange
    If yearChange >= 0 Then
        ' cell color to green
        Cells(printRow, 10).Interior.ColorIndex = 4
    Else
        ' cell color to red
        Cells(printRow, 10).Interior.ColorIndex = 3
    End If
    percentChange = Round((yearChange / openPrice) * 100, 2)
    Cells(printRow, 11).Value = percentChange
    Cells(printRow, 12).Value = totalVolume
End Sub

' This subroutine creates all the headers needed
Sub PrintHeaders()
    ' Insert Headers for colated data
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    ' Insert Headers for award data
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
End Sub
