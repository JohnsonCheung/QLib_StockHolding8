Attribute VB_Name = "MxIdeMthDup"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthDup."
Public Const FFoDupMth$ = MthcFF

Private Sub MthDupnyP__Tst()
VcAy MthDupnyP(True, True)
End Sub

Function MthDupnyP(Optional InlPrv As Boolean, Optional IsExactDup As Boolean) As String()  ' Return Mthn.Mdn[.Mdy]
Dim A$(): A = Mth4nyP
Dim Dy(): Dy = W1Dy(A)
Dim DyPrv(): DyPrv = W1WhPrv(Dy, InlPrv) ' Rmv Prv if need
Dim N$(): N = W1Mthny(DyPrv)    ' Mthny
Dim Dup$(): Dup = AwDup(N)        ' Dup mthny
Dim DupI&(): DupI = W1Ixy(N, Dup) ' Ixy for each N which is those N is an element of DupMthny
Dim DyDup(): DyDup = AwIxy(DyPrv, DupI) ' Wh only Dup
Dim DyExa(): DyExa = W1WhExa(DyDup, Dup, IsExactDup)
MthDupnyP = QSrt(W1Dupny(DyExa))     ' Chg Dupn
End Function

Private Function W1Ixy(Mthny$(), Dup$()) As Long()
Dim J&: For J = 0 To UB(Mthny)
    If HasEle(Dup, Mthny(J)) Then PushI W1Ixy, J
Next
End Function

Private Function W1Dy(Mth4ny$()) As Variant()
Dim N: For Each N In Itr(Mth4ny)
    PushI W1Dy, SplitDot(N)
Next
End Function

Private Function W1WhPrv(Dy(), InlPrv As Boolean) As Variant()
If InlPrv Then W1WhPrv = Dy: Exit Function
Dim Dr: For Each Dr In Itr(Dy)
    If Si(Dr) = 3 Then PushI W1WhPrv, Dr ' Mth4n will have 3 or 4 seg.  For Pub, always has only 3 segment.  So only take this
Next
End Function

Private Function W1Mthny(Dy()) As String()
Dim Dr: For Each Dr In Itr(Dy)
    PushI W1Mthny, Dr(1) ' 2nd Seg will be mthn for a Mth4n
Next
End Function

Private Function W1WhExa(Dy(), Dup$(), IsExactDup As Boolean) As Variant() ' Return Subset of @Dy::Mth4ny which is @IsExactDup with help of @Dup::Mthny
If Not IsExactDup Then W1WhExa = Dy: Exit Function
Dim D: For Each D In Itr(Dup)     ' D::One-mthn
    PushIAy W1WhExa, W1ExaOne(Dy, D)
Next
End Function

Private Function W1ExaOne(Dy(), Mthn) As Variant() ' Return subset of Mth4nDy which having Dup and there 2 or more Dr are exact dup
Dim DyMth(): DyMth = W1WhDy(Dy, Mthn) ' Dy has all Mth4n which eq to mthn-@Dup
Dim Mthl$(): Mthl = W1Mthly(DyMth)           ' All Mthl with same mthn.
Dim ExaI&(): ExaI = W1ExaI(Mthl)          ' Ixy of *L with exact eq
W1ExaOne = AwIxy(DyMth, ExaI)
End Function

Private Function W1Mthly(Dy()) As String() ' Mthly of @Dy.  Return Mthl to each each Dr of @Mth4nDy
Dim Mdn$, Mthn$, ShtMthTy$
Dim Dr: For Each Dr In Itr(Dy)
    AsgAy Dr, Mdn, Mthn, ShtMthTy
    PushI W1Mthly, MthlzMN(Md(Mdn), Mthn, ShtMthTy)
Next
End Function

Private Function W1ExaI(Mthly$()) As Long() ' Exa Ixy. return Ixy of @Mthly which is same value each other
Dim O&()
Dim J&: For J = 0 To UB(Mthly)
    If Not HasEle(O, J) Then
        PushIAy O, W1ExaI1(J, Mthly)
    End If
Next
W1ExaI = O
End Function

Private Function W1ExaI1(Ix&, Mthly$()) As Long() ' Exa Ixy 1 itm.  which is Ixy for Mthly, which are exactly eq to the item pointed by @Ix&-of-@Mthly
Dim L$: L = Mthly(Ix)       ' The Mthly
Dim O&()                    ' to be oupt
Dim I&: For I = 0 To UB(Mthly)
    If I <> Ix Then
        If L = Mthly(I) Then
            PushI O, I
        End If
    End If
Next
If Si(O) > 0 Then PushI O, Ix ' Some of Mthly has same value of *L, *O will have ele, then Ix need to put to O
W1ExaI1 = O
End Function

Private Function W1WhDy(Dy(), Mthn) As Variant() ' Where @Dy.  which is Subset of @Dy::Mth4nDy having mthn = @Mthn
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(1) = Mthn Then PushI W1WhDy, Dr ' ele-1 is mthn for Mth4n
Next
End Function

Private Function W1Dupny(Dy()) As String()
Dim Dr: For Each Dr In Itr(Dy)
    PushI W1Dupny, W1Dupn(Dr)
Next
End Function

Private Function W1Dupn$(A) ' @A is Mth4n segments: Mdn.Mthn.ShtMthTy[.ShtMdy]
Select Case Si(A)
Case 3: W1Dupn = JnDotApNB(A(1), A(0), A(2))
Case 4: W1Dupn = JnDotApNB(A(1), A(0), A(2), A(3))
Case Else: Thw CSub, "Si of Mth4n-Dr must be 4 or 3", "But-now", Si(A)
End Select
End Function
Function DupMthny(Optional InlPrv As Boolean) As String()
Dim O$(), All$()
Dim C As VBComponent: For Each C In CPj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        Dim N3() As Mthn3: N3 = Mthn3yWhInlPrv(Mthn3yzM(C.CodeModule), InlPrv)
        Dim N$(): N = Mthnyz3(N3)
        Dim Mthn: For Each Mthn In Itr(N)
            If HasEle(All, Mthn) Then
                PushI O, Mthn
            Else
                PushI All, Mthn
            End If
        Next
    End If
Next
DupMthny = O
End Function

Function DupMthlny(Optional InlPrv As Boolean, Optional IsExactDup As Boolean) As String()
Dim D As Drs: D = DupMthDrs
Dim Md$():   Md = AmAliR(StrCol(D, "Mdn"))
Dim Mth$(): Mth = AmAli(StrCol(D, "Mthn"))
Dim A$():     A = Ay2JnDot(Md, Mth)
Dim Mthln$(): Mthln = StrCol(D, "Mthln")
                   DupMthlny = Ay2JnSngQ(A, Mthln)
End Function

Function DupMthWs() As Worksheet
Set DupMthWs = DupMthWszP(CPj)
End Function

Private Sub DupMthDrs__Tst()
BrwDrs DupMthDrs(True)
End Sub

Function DupMthDrs(Optional InlPrv As Boolean, Optional IsExactDup As Boolean) As Drs
DupMthDrs = DupMthDrszP(CPj, InlPrv, IsExactDup)
End Function

Function DupMthDrszP(P As VBProject, Optional InlPrv As Boolean, Optional IsExactDup As Boolean) As Drs
Dim A As Drs: A = DwEq(MthcDrszP(P), "MdTy", "Std")
Dim A1 As Drs:
    If InlPrv Then
        A1 = A
    Else
        A1 = DwNe(A, "Mdy", "Prv")
    End If
Dim B As Drs: B = DwDup(A1, "Mthn")
Dim C As Drs: C = SrtDrs(B, "Mthn")
Dim D As Drs: D = AddColzValIdqCnt(C, "Mthl")
If IsExactDup Then
    DupMthDrszP = DwDup(D, "MthlId")
Else
    DupMthDrszP = D
End If
End Function

Function FmtDupMthWs(DupMthWs As Worksheet) As Worksheet
Dim Lo As ListObject: Set Lo = FstLo(DupMthWs)
SetLcWdt Lo, "MthL", 10
SetLcWrp Lo, "MthL", False
End Function

Function DupMthWszP(P As VBProject) As Worksheet
Set DupMthWszP = FmtDupMthWs(WszDrs(DupMthDrszP(P), "DupMth"))
End Function

Sub BrwDupMth(Optional InlPrv As Boolean, Optional IsExactDup As Boolean)
BrwDrs DupMthDrs(InlPrv, IsExactDup)
End Sub
