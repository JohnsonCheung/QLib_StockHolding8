Attribute VB_Name = "MxXlsRgVal"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsRgVal."
Function FstDrzRg(A As Range) As Variant()
FstDrzRg = FstDrzSq(SqzRg(RgR(A, 1)))
End Function
Function FstColzRg(A As Range) As Variant()
FstColzRg = FstColzSq(SqzRg(RgC(A, 1).Value))
End Function
