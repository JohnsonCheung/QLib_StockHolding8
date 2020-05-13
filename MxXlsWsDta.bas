Attribute VB_Name = "MxXlsWsDta"
Option Explicit
Option Compare Text
Const CNs$ = "WsDta"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsWsDta."

Function DtaRg(S As Worksheet) As Range
Set DtaRg = S.Range(S.Cells(1, 1), LasCell(S))
End Function

Function DtaSq(S As Worksheet) As Variant()
DtaSq = DtaRg(S).Value
End Function

Function DtaDrs(S As Worksheet) As Drs
DtaDrs = DrszSq(DtaSq(S))
End Function
