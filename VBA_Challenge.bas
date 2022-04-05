Attribute VB_Name = "Module11"
Sub OrganizeTicker()
Dim sheet_num As Integer
Dim last_row As Long
Dim range_num As Range
Dim ticker As Long
Dim wb As Workbook
Dim ws As Worksheet
Dim stock_open As Double
Dim stock_close As Double
Dim last_ticker As Long
Dim ticker_num As Range
Dim stock_open_roam As Double
Dim value_found As Boolean
Dim greatest_increase_ticker As String
Dim greatest_increase_percent As Double
Dim greatest_decrease_ticker As String
Dim greatest_decrease_percent As Double
Dim greatest_volume_ticker As String
Dim greatest_volume_total As Variant



sheet_num = Application.Worksheets.Count





For i = 1 To sheet_num
    Worksheets(i).Activate
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(i)
    Set range_num = ws.Range("I1")
    Set ticker_num = ws.Range("A1")

    stock_open = 0
    stock_close = 0
    greatest_increase_ticker = "0"
    greatest_increase_percent = 0
    greatest_decrease_ticker = "0"
    greatest_decrease_percent = 0
    greatest_volume_ticker = "0"
    greatest_volume_total = 0
    
    ws.Range("A1").Value = "Ticker"
    
    ws.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=range_num, Unique:=True
    
    last_row = range_num.End(xlDown).Row
    last_ticker = ticker_num.End(xlDown).Row
    
    For x = 2 To last_row
        ws.Cells(x, 12) = WorksheetFunction.SumIf(ws.Range("A:A"), ws.Cells(x, 9), ws.Range("G:G"))
        
        value_found = False
        
        For j = 2 To last_ticker
            If (ws.Cells(x, 9).Value = ws.Cells(j, 1).Value) Then
                stock_close = ws.Cells(j, 6).Value
                If (ws.Cells(x, 9).Value = ws.Cells(j, 1).Value) And value_found = False Then
                    stock_open = ws.Cells(j, 3).Value
                    value_found = True
                End If
            End If
        Next j
        
        ws.Cells(x, 10) = stock_close - stock_open
        ws.Cells(x, 11) = (stock_close - stock_open) / stock_open
        ws.Cells(x, 11).NumberFormat = "0.00%"
        
        If ws.Cells(x, 10).Value > 0 Then
            ws.Cells(x, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(x, 10).Value < 0 Then
            ws.Cells(x, 10).Interior.ColorIndex = 3
        End If
        
        
        If ws.Cells(x, 11).Value > greatest_increase_percent Then
            greatest_increase_ticker = ws.Cells(x, 9).Value
            greatest_increase_percent = ws.Cells(x, 11).Value
        End If
        If ws.Cells(x, 11).Value < greatest_decrease_percent Then
            greatest_decrease_ticker = ws.Cells(x, 9).Value
            greatest_decrease_percent = ws.Cells(x, 11).Value
        End If
        If (ws.Cells(x, 12).Value / 10000) > greatest_volume_total Then
            greatest_volume_ticker = ws.Cells(x, 9).Value
            greatest_volume_total = ws.Cells(x, 12).Value / 10000
        End If
        
        
    Next x
    
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greastest Total Volume"
    ws.Range("P2").Value = greatest_increase_ticker
    ws.Range("Q2").Value = greatest_increase_percent
    ws.Range("P3").Value = greatest_decrease_ticker
    ws.Range("Q3").Value = greatest_decrease_percent
    ws.Range("P4").Value = greatest_volume_ticker
    ws.Range("Q4").Value = greatest_volume_total * 10000
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    
    
    
Next i



End Sub
