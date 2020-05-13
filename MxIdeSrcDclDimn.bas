Attribute VB_Name = "MxIdeSrcDclDimn"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Dim"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcDclDimn."

Function Dimn$(Ln)
Dim L$: L = Ln
If ShfTermX(L, "Dim") Then Dimn = TakNm(LTrim(L))
End Function

Function DimNy(Ly$()) As String()
Dim L
For Each L In Itr(Ly)
    PushI DimNy, Dimn(L)
Next
End Function
