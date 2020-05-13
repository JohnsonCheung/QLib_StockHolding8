Attribute VB_Name = "MxVbAyLyzSSAy"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CNs$ = "Str.Ly"
Const CMod$ = CLib & "MxVbAyLyzSSAy."

Function DyzSSAy(SSAy$()) As Variant()
Dim Dr: For Each Dr In Itr(SSAy)
    PushI DyzSSAy, SyzSS(Dr)
Next
End Function

Function AliSSAy(SSAy$()) As String()
Dim Dy(): Dy = DyzSSAy(SSAy)
Dim Dy1(): Dy1 = AliDy(Dy)
AliSSAy = JnDy(Dy1)
End Function
