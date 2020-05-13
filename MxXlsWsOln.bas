Attribute VB_Name = "MxXlsWsOln"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxXlsWsOln."

Sub MiniWbOLvl(Wb As Workbook)
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    MiniWsOLvl Ws
Next
End Sub
Sub MiniWsOLvl(Ws As Worksheet)
Ws.Outline.ShowLevels 1, 1
End Sub
