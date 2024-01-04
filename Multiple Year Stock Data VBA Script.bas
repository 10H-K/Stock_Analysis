Attribute VB_Name = "MultipleYearStockDataVBA"
Sub MasterSwitch()

    'Call Sub For Raw Data Analysis

    StockRawDataAnalysis

    'Call Sub For Condensed Data Analysis

    StockCondensedDataAnalysis

    'Call Sub For Naming And Formatting

    NamingAndFormatting

    'Message Showing User Data Has Been Organized
    ThisWorkbook.Sheets("2018").Range("P16").Value = "Data Is Organized!"
    ThisWorkbook.Sheets("2018").Range("P16").Font.Name = "Times New Roman"
    ThisWorkbook.Sheets("2018").Range("P16").Font.Size = "12"

End Sub


Sub StockRawDataAnalysis()

    'Loop Sub Procedure Through Each Worksheet In Workbook
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets

        'Output Location For Summarized Stock Data
        Dim DataOutput As Integer
        DataOutput = 5

        'Loop Through Each Row Of Raw Stock Data
        Dim LastRow As Long
        Dim i As Long
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
        
            'Variables For Stock Ticker Identity & Cumulative Volume Traded
            Dim StockTicker As String
            Dim TotalVolume As Double

            'Set Initial Price In First Row After Stock Ticker Identity Change
            Dim OpenPrice As Double
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                OpenPrice = ws.Cells(i, 3).Value
            Else
            End If

            'Identification Of Stock Ticker Identity Change As Loop Progresses
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'Set Current Stock Ticker Before Identity Change
                StockTicker = ws.Cells(i, 1).Value

                'Set Close Price In Final Row Before Identity Change
                Dim ClosePrice As Double
                ClosePrice = ws.Cells(i, 6).Value
                
                'Yearly Price Change/Percent Change Calculations
                Dim PriceChange As Double
                Dim PercentChange As Double
                PriceChange = ClosePrice - OpenPrice
                PercentChange = ((ClosePrice - OpenPrice) / OpenPrice)

                'Add Volume Traded To Total Stock Volume Traded Before Identity Change
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value

                'Print Stock Ticker To Output Location
                ws.Range("J" & DataOutput).Value = StockTicker

                'Print Total Stock Volume Traded To Output Location
                ws.Range("M" & DataOutput).Value = TotalVolume

                'Print Yearly Change To Output Location
                ws.Range("K" & DataOutput).Value = PriceChange

                'Print Yearly Percent Change To Output Location
                ws.Range("L" & DataOutput).Value = PercentChange

                'Reset All Variables For Stock Ticker Identity Change
                StockTicker = ""
                TotalVolume = 0
                OpenPrice = 0
                ClosePrice = 0
                PriceChange = 0
                PercentChange = 0
                
                'Add Row To Output Location For Stock Ticker Identity Change
                DataOutput = DataOutput + 1

            'If Stock Ticker Identity Change Absent, Keep Accumulating Totals
            Else

                'Continue Adding To Total Stock Volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value

            End If

        Next i

    Next ws

End Sub


Sub StockCondensedDataAnalysis()

    'Loop Sub Procedure Through Each Worksheet In Workbook
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets

        'Output Location For Greatest Percent Increase/Percent Decrease/Total Volume
        Dim DataOutput As Integer
        DataOutput = 5

        'Variables For Greatest Percent Increase/Percent Decrease/Total Volume Determination
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double
        Dim CurrentPercent As Double
        Dim CurrentVolume As Double
        
        'Reset Variables For Greatest Percent Increase/Percent Decrease/Total Volume Determination
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        CurrentPercent = 0
        CurrentVolume = 0

        'Stock Tickers Corresponding To Greatest Percent Increase/Percent Decrease/Total Volume Values
        Dim StockTickerIncrease As String
        Dim StockTickerDecrease As String
        Dim StockTickerVolume As String
        Dim CurrentTicker As String
        
        'Reset Stock Tickers Corresponding To Greatest Percent Increase/Percent Decrease/Total Volume Values
        StockTickerIncrease = ""
        StockTickerDecrease = ""
        StockTickerVolume = ""
        CurrentTicker = ""

        'Loop Through Summarized Stock Data
        Dim LastRow As Long
        Dim i As Long
        LastRow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
        For i = 5 To LastRow
        
            'Assign Values To Variables Holding Current Values During Loop
            CurrentPercent = ws.Cells(i, 12).Value
            CurrentVolume = ws.Cells(i, 13).Value
            CurrentTicker = ws.Cells(i, 10).Value

            'Check For Greatest Percent Increase
            If CurrentPercent > GreatestIncrease Then
                GreatestIncrease = CurrentPercent
                StockTickerIncrease = CurrentTicker
            End If

            'Check For Greatest Percent Decrease
            If CurrentPercent < GreatestDecrease Then
                GreatestDecrease = CurrentPercent
                StockTickerDecrease = CurrentTicker
            End If

            'Check For Greatest Total Volume
            If CurrentVolume > GreatestVolume Then
                GreatestVolume = CurrentVolume
                StockTickerVolume = CurrentTicker
            End If

        Next i

        'Print Greatest Percent Increase To Output Location
        ws.Range("R" & DataOutput).Value = GreatestIncrease
        ws.Range("Q" & DataOutput).Value = StockTickerIncrease

        'Print Greatest Percent Decrease To Output Location
        DataOutput = DataOutput + 1
        ws.Range("R" & DataOutput).Value = GreatestDecrease
        ws.Range("Q" & DataOutput).Value = StockTickerDecrease

        'Print Greatest Total Volume To Output Location
        DataOutput = DataOutput + 1
        ws.Range("R" & DataOutput).Value = GreatestVolume
        ws.Range("Q" & DataOutput).Value = StockTickerVolume

    Next ws

End Sub


Sub NamingAndFormatting()

    'Loop Sub Procedure Through Each Worksheet In Workbook
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets

        'Define Variable
        Dim i As Long

        'Assign Names To Cells
        ws.Range("J4,Q4").Value = "Ticker"
        ws.Range("P4").Value = "Parameters"
        ws.Range("K4").Value = "Yearly Change"
        ws.Range("L4").Value = "Percent Change"
        ws.Range("M4").Value = "Total Stock Volume"
        ws.Range("R4").Value = "Value"
        ws.Range("P5").Value = "Greatest % Increase"
        ws.Range("P6").Value = "Greatest % Decrease"
        ws.Range("P7").Value = "Greatest Total Volume"

        'Assign Variables To Columns To Display Data Approriately
        Dim LastRow As Long
        Dim CurrencyColumn As Range
        Dim PercentageColumnA As Range
        Dim PercentageColumnB As Range
        Set CurrencyColumn = ws.Columns("K")
        Set PercentageColumnA = ws.Columns("L")
        Set PercentageColumnB = Union(ws.Range("R5"), ws.Range("R6"))

        'Table Insertion For Stock Highlight Data
        Dim Table As ListObject
        Set Table = ws.ListObjects.Add(xlSrcRange, ws.Range("P4:R7"), , xlYes)

        'Insertion Of Color Formatting For Yearly/Percent Change
        LastRow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
        For i = 5 To LastRow
            If ws.Cells(i, 11) > 0 Then
                ws.Cells(i, 11).Interior.Color = RGB(0, 255, 0) 'Green Color Yearly Change
                ws.Cells(i, 12).Interior.Color = RGB(0, 255, 0) 'Green Color Yearly Percent Change
            ElseIf ws.Cells(i, 11) < 0 Then
                ws.Cells(i, 11).Interior.Color = RGB(255, 0, 0) 'Red Color Yearly Change
                ws.Cells(i, 12).Interior.Color = RGB(255, 0, 0) 'Red Color Yearly Percent Change
            Else
                ws.Cells(i, 11).Interior.Color = RGB(255, 255, 0) 'Yellow Color Yearly Change
                ws.Cells(i, 12).Interior.Color = RGB(255, 255, 0) 'Yellow Color Yearly Percent Change
            End If

            'Format Data Appropriately Part 1
            CurrencyColumn.NumberFormat = "$#,##0.00"
            PercentageColumnA.NumberFormat = "0.00%"
            
        Next i
        
        'Format Data Appropriately Part 2
        PercentageColumnB.NumberFormat = "0.00%"

        'AutoAdjust Font and Font Size
        ws.Columns("A:G").Font.Name = "Times New Roman"
        ws.Columns("J:M").Font.Name = "Times New Roman"
        ws.Columns("P:R").Font.Name = "Times New Roman"
        ws.Range("A1:G1").Font.Bold = "True"
        ws.Range("J4:M4").Font.Bold = "True"
        ws.Range("P4:R4").Font.Bold = "True"
        ws.Columns("A:G").Font.Size = "12"
        ws.Columns("J:M").Font.Size = "12"
        ws.Columns("P:R").Font.Size = "12"

        'AutoFit Column Width and Center Text
        ws.Columns("J:M").AutoFit
        ws.Columns("P:R").ColumnWidth = 22
        ws.Columns("A:G").HorizontalAlignment = xlCenter
        ws.Columns("J:M").HorizontalAlignment = xlCenter
        ws.Columns("P:R").HorizontalAlignment = xlCenter

    Next ws

End Sub
