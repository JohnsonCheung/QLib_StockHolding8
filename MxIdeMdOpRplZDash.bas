Attribute VB_Name = "MxIdeMdOpRplZDash"
Option Compare Text
Option Explicit
Const CNs$ = "Mth.Ln.Rpl"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdRplZDash."
Sub RplZSub()
RplPj_ CPj
End Sub
Private Sub RplPj_(Pj As VBProject)
RplPjLNewO Pj, ZSubPjLNewOAy(Pj)
End Sub
Private Sub ZSubPjLNewO__Tst()
BrwPjLNewOAy ZSubPjLNewOAy(CPj)
End Sub
Private Function ZSubPjLNewOAy(Pj As VBProject) As PjLNewO()
Dim C As VBComponent: For Each C In Pj.VBComponents
    Dim D As Drs: D = ZSubLNewO(C.CodeModule)
    If Not NoReczDrs(D) Then
        PushPjLNewO ZSubPjLNewOAy, PjLNewO(C.Name, D)
    End If
Next
End Function
Private Function ZSubLNewO(M As CodeModule) As Drs
ZSubLNewO = ZSubLNewOzSrc(Src(M))
End Function
Private Function ZSubLNewOzSrc(Src$()) As Drs
Dim Dy(), Lno&
Dim L: For Each L In Itr(Src)
    Lno = Lno + 1
    If IsZSubLin(L) Then PushI Dy, Array(Lno, RplZSubLin(L), L)
Next
ZSubLNewOzSrc = Drs(SyzSS(LNewOFF), Dy)
End Function

Private Function IsZSubLin(L) As Boolean
Dim Ln$: Ln = RmvMdy(L)
If ShfMthTy(Ln) <> "Sub" Then Exit Function
IsZSubLin = HasPfx(Ln, "Z_")
End Function

Private Function RplZSubLin$(ZSubLin)
Dim N$: N = Mthn(ZSubLin)
If Fst2Chr(N) <> "Z_" Then Thw CSub, "Given @ZSubLin is not ZSub-Mth-Ln", "ZSubLin", ZSubLin
RplZSubLin = FmtQQ("Private Sub ?__Tst()", RmvFst2Chr(N))
End Function
