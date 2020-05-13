Attribute VB_Name = "MxXlsWsSetWbOutLinSum"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsWsSetWbOutLinSum."

Sub SetWbOutLinSum(Wb As Workbook)
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    SetWsOutLinSum Ws
Next
End Sub
Sub SetWsOutLinSum(Ws As Worksheet)
SetWsSumRow Ws
SetWsSumCol Ws
End Sub
Sub SetWbSumCol(Wb As Workbook, Optional SumCol As XlSummaryColumn = XlSummaryColumn.xlSummaryOnLeft)
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    SetWsSumCol Ws, SumCol
Next
End Sub
Sub SetWsSumCol(Ws As Worksheet, Optional SumCol As XlSummaryColumn = XlSummaryColumn.xlSummaryOnLeft)
On Error Resume Next
Ws.Outline.SummaryColumn = SumCol
End Sub

Sub SetWbSumRow(Wb As Workbook, Optional SumRow As XlSummaryRow = XlSummaryRow.xlSummaryAbove)
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    SetWsSumRow Ws, SumRow
Next
End Sub
Sub SetWsSumRow(Ws As Worksheet, Optional SumRow = XlSummaryRow.xlSummaryAbove)
On Error Resume Next
Ws.Outline.SummaryRow = SumRow
End Sub
