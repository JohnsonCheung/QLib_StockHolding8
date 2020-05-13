Attribute VB_Name = "MxVbStrBrkLin"
Option Explicit
Option Compare Text
Const CNs$ = "MxLinesInf"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxVbStrBrkLin."

Function LnBrkDy(Ly$(), Sep$()) As Variant()
'Ret : :Dy ! a dry wi each rec as a sy of brkg one Ln of @Ly.  Each Ln is brk by @Sep using fun-BrkLn @@
Dim L: For Each L In Itr(Ly)
    PushI LnBrkDy, BrkLn(L, Sep)
Next
End Function

Function BrkLn(Ln, Sep$(), Optional IsRmvSep As Boolean) As String()
'Ret : seg ay of a Ln sep by @Sep.  Si of seg ret = si of @sep + 1.  Each will have its own sep, expt fst.
'      Segs are not trim and wi/wo by @IsRmvSep.  If not @IsRmvSep, Jn(@Rslt) will eq @Ln @@
Dim L$: L = Ln
Dim O$()
Dim S: For Each S In Sep
    PushI O, ShfBef(L, CStr(S))
Next
PushI O, L
If IsRmvSep Then
    Dim J&, Seg: For Each Seg In O
        PushI BrkLn, RmvPfx(Seg, Sep(J))
        J = J + 1
    Next
Else
    BrkLn = O
End If
End Function

Function LnBrkDyzSS(Ly$(), SepSS$) As Variant()
LnBrkDyzSS = LnBrkDy(Sy, SyzSS(SepSS))
End Function

Private Sub LnBrkDyzSS__Tst()
Dim Ly$(), Sep$
GoSub T0
Exit Sub
T0:
    Sep = ". . . . . ."
    Ly = Sy("AStkShpCst_Rpt.OupFx.Fun.")
    Ept = Sy("AStkShpCst_Rpt", ".OupFx", ".Fun", ".")
    GoTo Tst
Tst:
    Act = LnBrkDyzSS(Sy, Sep)
    C
    Return
End Sub
