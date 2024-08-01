Attribute VB_Name = "Module1"
Option Explicit

Sub Stock_Macro()

' Declare variables
Dim ws As Worksheet
Dim lastrow As Long
Dim i As Double
Dim sum_table_row As Long
Dim Ticker As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim qtr_Change As Double
Dim Percent_Change As Double
Dim tot_Vol As Double
Dim Gratest_percent_increase As Double
Dim Gratest_percent_decrease As Double
Dim Highest_Volume As Double
Dim UpTicker As String
Dim DownTicker As String
Dim VolumeTicker As String

' Loop through all worksheets in the workbook
For Each ws In ThisWorkbook.Worksheets
    ' Initialize variables for each sheet
    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    sum_table_row = 2
    Gratest_percent_increase = 0
    Gratest_percent_decrease = 0
    Highest_Volume = 0

    ' Set up summary table headers in each sheet
    ws.Range("I1") = "Ticker Symbol"
    ws.Range("J1") = "Quarterly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"

    ' Loop through each row in the sheet
    For i = 2 To lastrow
        ' Check if ticker symbol has changed
        If ws.Cells(i, "A").Value <> ws.Cells(i - 1, "A").Value Then
            Ticker = ws.Cells(i, "A").Value
            OpenPrice = ws.Cells(i, "C").Value
            tot_Vol = 0
        End If

        ' Add current row's volume to total volume
        tot_Vol = tot_Vol + ws.Cells(i, "G").Value

        ' Check if ticker symbol's last row
        If ws.Cells(i, "A").Value <> ws.Cells(i + 1, "A").Value Then
            ClosePrice = ws.Cells(i, "F").Value
            qtr_Change = ClosePrice - OpenPrice
            If OpenPrice <> 0 Then
                Percent_Change = qtr_Change / OpenPrice
            Else
                Percent_Change = 0
            End If

            ' Write values to the summary table
            ws.Cells(sum_table_row, "I").Value = Ticker
            ws.Cells(sum_table_row, "J").Value = qtr_Change
            ws.Cells(sum_table_row, "K").Value = Percent_Change
            ws.Cells(sum_table_row, "L").Value = tot_Vol

            ' Color code the quarterly change cell
            If qtr_Change >= 0 Then
                ws.Cells(sum_table_row, "J").Interior.ColorIndex = 4
            Else
                ws.Cells(sum_table_row, "J").Interior.ColorIndex = 3
            End If

            ' Check for greatest percent increase/decrease and highest volume
            If Percent_Change > Gratest_percent_increase Then
                Gratest_percent_increase = Percent_Change
                UpTicker = Ticker
            ElseIf Percent_Change < Gratest_percent_decrease Then
                Gratest_percent_decrease = Percent_Change
                DownTicker = Ticker
            End If

            If tot_Vol > Highest_Volume Then
                Highest_Volume = tot_Vol
                VolumeTicker = Ticker
            End If

            ' Move to the next row in the summary table
            sum_table_row = sum_table_row + 1
        End If
    Next i

    ' Create table for greatest increase, decrease, and total volume
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"

    ws.Range("O1").Value = "Ticker"
    ws.Range("O2").Value = UpTicker
    ws.Range("O3").Value = DownTicker
    ws.Range("O4").Value = VolumeTicker

    ws.Range("P1").Value = "Value"
    ws.Range("P2").Value = Gratest_percent_increase
    ws.Range("P3").Value = Gratest_percent_decrease
    ws.Range("P4").Value = Highest_Volume

    ' Format cells for percentage and volume
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("P2:P3").NumberFormat = "0.00%"
    ws.Range("P4").NumberFormat = "#,##0"
    ws.Range("L:L").NumberFormat = "#,##0"

    ' Autofit columns for better visibility
    ws.Columns("I:P").AutoFit
Next ws

End Sub

