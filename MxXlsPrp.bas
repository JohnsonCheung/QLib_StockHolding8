Attribute VB_Name = "MxXlsPrp"
Option Explicit
Option Compare Text
Const CNs$ = "Xls"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsPrp."
Function PjnyzXls(X As Excel.Application) As String()
PjnyzXls = PjnyzV(X.Vbe)
End Function
