Attribute VB_Name = "MxVbFmtTaby"
Option Compare Text
Option Explicit
Const CNs$ = "Fmt.Taby"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxVbFmtTaby."
Function DyzTsy(Tsy$()) As Variant()
Dim L: For Each L In Itr(Tsy)
    PushI DyzTsy, SplitTab(L)
Next
End Function
Function FmtTsy(Tsy$()) As String()
FmtTsy = FmtDy(DyzTsy(Tsy))
End Function
