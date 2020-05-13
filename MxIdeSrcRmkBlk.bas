Attribute VB_Name = "MxIdeSrcRmkBlk"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Rmk"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcRmkBlk."

Private Sub Vrmk__Tst()
VcLyAy Vrmk(SrczP(CPj))
End Sub

Function Vrmk(Src$()) As Variant()
':Vrmk: LyAy
Dim O$(), InBlk As Boolean, IsRmk As Boolean
Dim L: For Each L In Itr(Src)
    IsRmk = IsVrmkLn(L)
    Select Case True
    Case InBlk And IsRmk
        PushI O, L
    Case InBlk
        If SomEle(O) Then PushI Vrmk, O
        InBlk = False
    Case IsRmk
        If SomEle(O) Then PushI Vrmk, O
        InBlk = True
        O = Sy(L)
    End Select
Next
If SomEle(O) Then PushI Vrmk, O
End Function
