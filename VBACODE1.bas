Attribute VB_Name = "Module1"
Sub AnalyzeStockDataForSpecificQuarters()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim totalVolume As Double
    Dim percentageChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String
    Dim f As Long
    Dim r As Integer
    Dim i As Long
    Dim quarterSheets As Variant
    Dim sheetName As Variant

    
    quarterSheets = Array("Q1", "Q2", "Q3", "Q4")

    For Each sheetName In quarterSheets
        Set ws = ThisWorkbook.Sheets(sheetName)

        greatestIncrease = -99999
        greatestDecrease = 99999
        greatestVolume = 0

        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        If ws.Cells(1, 9).Value <> "Ticker" Then
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Total Volume"
        End If

        r = 2
        f = 2
        totalVolume = 0
        
        For i = 2 To lastRow


            ticker = ws.Cells(i, 1).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            closePrice = ws.Cells(i, 6).Value


            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastRow Then

                openPrice = ws.Cells(f, 3).Value
                quarterlyChange = closePrice - openPrice
                
                If openPrice <> 0 Then
                    percentageChange = (quarterlyChange / openPrice)
                Else
                    percentageChange = 0
                End If

                
                ws.Cells(r, 9).Value = ticker
                ws.Cells(r, 10).Value = quarterlyChange
                ws.Cells(r, 11).Value = percentageChange
                ws.Cells(r, 12).Value = totalVolume
                ws.Cells(r, 11).NumberFormat = "0.00%"

            
                With ws.Cells(r, 10).FormatConditions
                    .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
                    .Item(1).Interior.Color = RGB(0, 255, 0)

                    .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
                    .Item(2).Interior.Color = RGB(255, 0, 0)
                End With

                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    increaseTicker = ticker
                End If

                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    decreaseTicker = ticker
                End If

            
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTicker = ticker
                End If

    
                r = r + 1
                totalVolume = 0
                f = i + 1
            End If
        Next i

        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = increaseTicker
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"

        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = decreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"

        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = volumeTicker
        ws.Cells(4, 17).Value = greatestVolume
        ws.Cells(4, 17).NumberFormat = "0.00E+00"
    Next sheetName

End Sub

