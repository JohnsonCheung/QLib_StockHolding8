Attribute VB_Name = "MxXlsLoDta"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsLoDta."
Function DrzLoCell(Lo As ListObject, Cell As Range) As Variant()
Dim Ix&: Ix = LoRno(Lo, Cell): If Ix = -1 Then Exit Function
DrzLoCell = FstDrzRg(Lo.ListRows(Ix).Range)
End Function
