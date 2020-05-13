Attribute VB_Name = "MxXlsCur"
Option Explicit
Option Compare Text
Const CNs$ = "Cur.Xls"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsCur."
Function CWs() As Worksheet
Set CWs = Xls.ActiveSheet
End Function

Function CWb() As Workbook
Set CWb = Xls.ActiveWorkbook
End Function

Function Xls() As Excel.Application
Set Xls = Excel.Application
End Function
