Attribute VB_Name = "MxXlsWsSetNoAutoColWdt"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsWsSetNoAutoColWdt."

Sub SetWbNoAutoColWdt(Wb As Workbook)
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    SetWsNoAutoColWdt Ws
Next
End Sub

Sub SetWsNoAutoColWdt(Ws As Worksheet)
Dim Lo As ListObject: For Each Lo In Ws.ListObjects
    SetLoNoAutoColWdt Lo
Next
End Sub
Sub SetLoNoAutoColWdt(Lo As ListObject)
Lo.QueryTable.AdjustColumnWidth = False
End Sub
