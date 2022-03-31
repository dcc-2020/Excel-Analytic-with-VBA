Attribute VB_Name = "Module1"
Sub OrganizeTicker()
Dim sheet_num As Integer
Dim last_row As Long
Dim range_num As Range
Dim ticker As Long
Dim wb As Workbook
Dim ws As Worksheet



sheet_num = Application.Worksheets.Count


For i = 1 To sheet_num
    Worksheets(i).Activate
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(i)
    Set range_num = ws.Range("I1")
    
    ws.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=range_num, Unique:=True
    
    last_row = range_num.End(xlDown).Row
    For x = 2 To last_row
        ws.Cells(x, 10) = WorksheetFunction.SumIf(ws.Range("A:A"), ws.Cells(x, 9), ws.Range("L:L"))
    Next x
Next i




End Sub
