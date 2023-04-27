Attribute VB_Name = "Module1"
Sub StockData()

'Variable Declarations for Script
Dim Ticker As String
Dim DataSummary As Integer
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim TickerChange As Long
Dim ws As Worksheet
Dim GreatestIncrease As Double
Dim GreatestIncreaseTIcker As String
Dim GreatestDecrease As Double
Dim GreatestDecreaseTicker As String
Dim GreatestVolume As Double
Dim GreatestVolumeTicker As String

DataSummary = 2
YearlyChange = 0



'Format Loop to Run on all Sheets
For Each ws In Worksheets

TickerChange = 2

    'Insert Titles for Created Columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Insert Ticker Names into Ticker Column
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value



        'Calculate Total Stock Volume of each Ticker
        TotalStockVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(TickerChange, 7), ws.Cells(i, 7)))

        'Calculate the Yearly Change & Percent of each stock
        OpenPrice = ws.Cells(TickerChange, 3).Value
        ClosePrice = ws.Cells(i, 6).Value
        TickerChange = (i + 1)


        'Yearly Change Calculation Formulas
        YearlyChange = ClosePrice - OpenPrice
        PercentChange = YearlyChange / OpenPrice
        
        If PercentChange > GreatestIncrease Then
            GreatestIncrease = PercentChange
            GreatestIncreaseTIcker = Ticker
            
        End If
            
        If PercentChange < GreatestDecrease Then
            GreatestDecrease = PercentChange
            GreatestDecreaseTicker = Ticker
        End If
        
        If TotalStockVolume > GreatestVolume Then
            GreatestVolume = TotalStockVolume
            GreatestVolumeTicker = Ticker
        End If
            
        
        
        
        
        
        'Format numbers correctly in percentages & whole number format
        ws.Range("K" & DataSummary).NumberFormat = "0.00%"
        ws.Range("J" & DataSummary).NumberFormat = "0.00"


        'Insert Values into correct columns
        ws.Range("J" & DataSummary).Value = YearlyChange
        ws.Range("K" & DataSummary).Value = PercentChange
        ws.Range("I" & DataSummary).Value = Ticker
        ws.Range("L" & DataSummary).Value = TotalStockVolume



        'Conditional Formatting for -/+ Yearly Change
        If ws.Range("J" & DataSummary).Value > 0 Then
            ws.Range("J" & DataSummary).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & DataSummary).Value < 0 Then
            ws.Range("J" & DataSummary).Interior.ColorIndex = 3
        End If
        'Conditional Formatting for +/- Percent Change
        If ws.Range("K" & DataSummary).Value > 0 Then
            ws.Range("K" & DataSummary).Interior.ColorIndex = 4
        ElseIf ws.Range("K" & DataSummary).Value < 0 Then
            ws.Range("K" & DataSummary).Interior.ColorIndex = 3
        End If



    'Insert Titles for Greatest Columns
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"


    
    'Insert Values for Greatest Columns
    ws.Range("Q2").Value = GreatestIncrease
    ws.Range("Q3").Value = GreatestDecrease
    ws.Range("Q4").Value = GreatestVolume
    ws.Range("P2").Value = GreatestIncreaseTIcker
    ws.Range("P3").Value = GreatestDecreaseTicker
    ws.Range("P4").Value = GreatestVolumeTicker
  

    
    

    DataSummary = DataSummary + 1


    End If
Next i

DataSummary = 2
YearlyChange = 0
GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0


Next ws


End Sub
