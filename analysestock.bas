Attribute VB_Name = "Module1"
Sub AnalyseStocks()
    Dim row As Long
    Dim nextRow As Long
    Dim totalStockVolume As Double
    Dim ws As Worksheet
    Dim opening As Double
    Dim closing As Double
    Dim rowCount As Long
    Dim greatestTotalStockVolumeTicker As String
    Dim greatestTotalStockVolumeValue As Double
    Dim greatestIncreaseTicker As String
    Dim greatestIncreaseValue As Double
    Dim greatestDecreaseTicker As String
    Dim greatestDecreaseValue As Double
    
    
    
    'Set ws = ActiveSheet
For Each ws In Sheets
    
    totalStockVolume = 0
    greatestIncreaseValue = 0
    greatestIncreaseTicker = ""
    greatestDecreaseValue = 0
    greatestDecreaseTicker = ""
    greatestTotalStockVolumeTicker = ""
    greatestTotalStockVolumeValue = 0
    nextRow = 2
    
    
    'Create column headers
    ws.Range("I1") = "Tickers"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    
    ' Get the number of Rows in the worksheet
    rowCount = ws.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).row
    
    'Loop through all tickers
    For row = 2 To rowCount
    
        ' Check if we are begining a ticker
        If ws.Cells(row - 1, 1).Value <> ws.Cells(row, 1).Value Then
            totalStockVolume = 0
            ' save the opening price
             opening = ws.Cells(row, 3).Value
            
        End If
        
        ' For each row add the amount of volume to total
        totalStockVolume = totalStockVolume + ws.Cells(row, 7).Value
        
        
        If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            ws.Cells(nextRow, 9).Value = ws.Cells(row, 1).Value
            ' Save the closing price
            closing = ws.Cells(row, 6).Value
            ' Compute changes and store them into the spreadsheet
             ws.Cells(nextRow, 10).Value = closing - opening
             ws.Cells(nextRow, 11).Value = FormatPercent((closing - opening) / opening)
            ' Color the Yearly Change
                If ws.Cells(nextRow, 10).Value >= 0 Then
                    ws.Cells(nextRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(nextRow, 10).Interior.ColorIndex = 3
                End If
            ws.Cells(nextRow, 12).Value = totalStockVolume
            
            ' if totalStockVolume is greater than greatestTotalStockVolumeValue
            ' Then update greatestTotalStockVolumeValue
            ' and greatestTotalStockVolumeTicker
                If greatestIncreaseValue > ws.Cells(nextRow, 11).Value Then
                    greatestIncreaseValue = ws.Cells(nextRow, 11).Value
                    greatestIncreaseTicker = ws.Cells(nextRow, 9).Value
                End If
                If greatestDecreaseValue < ws.Cells(nextRow, 11).Value Then
                    greatestDecreaseValue = ws.Cells(nextRow, 11).Value
                    greatestDecreaseTicker = ws.Cells(nextRow, 9).Value
                End If
                If greatestTotalStockVolumeValue < totalStockVolume Then
                    greatestTotalStockVolumeValue = totalStockVolume
                    greatestTotalStockVolumeTicker = ws.Cells(nextRow, 9).Value
                End If
                
                
                
            nextRow = nextRow + 1
        End If
    Next row
    
    'Create secondary table
    'Create Column headers
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Cells(4, 17).Value = greatestTotalStockVolumeValue
    ws.Cells(4, 16).Value = greatestTotalStockVolumeTicker
    ws.Cells(3, 16).Value = greatestIncreaseTicker
    ws.Cells(3, 17).Value = FormatPercent(greatestIncreaseValue)
    ws.Cells(2, 16).Value = greatestDecreaseTicker
    ws.Cells(2, 17).Value = FormatPercent(greatestDecreaseValue)
    
Next ws
    
    
End Sub


