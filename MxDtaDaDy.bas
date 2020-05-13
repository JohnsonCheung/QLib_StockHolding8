Attribute VB_Name = "MxDtaDaDy"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaDy."

Function JnDotDy(Dy()) As String()
Dim Dr: For Each Dr In Itr(Dy)
    PushI JnDotDy, JnDot(Dr)
Next
End Function

Function AliRzDyC(Dy(), C) As Variant()
Dim Ay$(): Ay = AmAliR(ColzDy(Dy, C))
Dim O(): O = Dy
Dim J&
For J = 0 To UB(O)
    O(J)(C) = Ay(J)
Next
AliRzDyC = O
End Function

Function CntDyWhGt1(CntDy()) As Variant()
Dim Dr
For Each Dr In CntDy
    If Dr(1) > 1 Then PushI CntDyWhGt1, Dr
Next
End Function

Function CntDizDyC(Dy(), C&) As Dictionary
Set CntDizDyC = CntDi(ColzDy(Dy, C))
End Function

Function ColyzDy(Dy()) As Variant()
Dim J%: For J = 0 To NColzDy(Dy)
    PushI ColyzDy, ColzDy(Dy, J)
Next
End Function

Function DisCol(A As Drs, C$)
DisCol = AwDis(ColzDy(A.Dy, IxzAy(A.Fny, C)))
End Function

Function DisStrCol(A As Drs, C$) As String()
Dim I%: I = IxzAy(A.Fny, C)
Dim Col$(): Col = StrColzDy(A.Dy, I)
DisStrCol = AwDis(Col)
End Function

Function DistColzDy(Dy(), C&) As Variant()
DistColzDy = AwDis(ColzDy(Dy, C))
End Function

Function DotSyzDy(Dy()) As String()
Dim Dr
For Each Dr In Itr(Dy)
    PushI DotSyzDy, JnDot(Dr)
Next
End Function

Function DyJnFldNFld(Dy(), FstNFld%, Optional Sep$ = " ") As Variant()
Dim U%: U = FstNFld - 1
Dim UK%: UK = U - 1
Dim O(), Dr
For Each Dr In Itr(Dy)
    If U <> UB(Dr) Then
        ReDim Preserve Dr(U)
    End If
    Dim Ix: Ix = RowIxOptzDyDr(O, FstNEle(Dr, UK))
    If Ix = -1 Then
        PushI O, Dr
    Else
        Stop
'        O(Ix)(U) = AddNB(O(Ix)(U), Sep) & Dr(U)
    End If
Next
DyJnFldNFld = O
End Function

Function DyJnFldKK(Dy(), KKIxy&(), JnFldIx&, Optional Sep$ = " ") As Variant()
'Ret : :Dy-@KKIxy-@JnFldIx ! Ret Dy of Si(@KKIxy) + 1 columns with UKey-KKIxy
Dim Ixy&(): Ixy = KKIxy: PushI Ixy, JnFldIx
Dim N%: N = Si(Ixy)
DyJnFldKK = DyJnFldNFld(SelCol(Dy, Ixy), N)
End Function

Function DyoSq(Sq()) As Variant()
If Si(Sq) = 0 Then Exit Function
Dim R&: For R = 1 To UBound(Sq(), 1)
    PushI DyoSq, DrzSqr(Sq, R)
Next
End Function

Function DyoSslAy(SslAy$()) As Variant()
Dim Ssl$, I
For Each I In Itr(SslAy)
    Ssl = I
    PushI DyoSslAy, SyzSS(Ssl)
Next
End Function

Function DywCCNe(Dy(), C1&, C2&) As Variant()
Dim Dr
For Each Dr In Dy
    If Dr(C1) <> Dr(C2) Then PushI DywCCNe, Dr
Next
End Function

Function DywColGt(Dy(), C%, GtV) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C) > GtV Then PushI DywColGt, Dr
Next
End Function

Function DywCoLiny(Dy(), ColIx%, InAy) As Variant()
Const CSub$ = CMod & "DywCoLiny"
If Not IsArray(InAy) Then Thw CSub, "[InAy] is not Array, but [TypeName]", "InAy-TypeName", TypeName(InAy)
If Si(InAy) = 0 Then DywCoLiny = Dy: Exit Function
Dim Dr
For Each Dr In Itr(Dy)
    If HasEle(InAy, Dr(ColIx)) Then PushI DywCoLiny, Dr
Next
End Function

Function DywColNe(Dy(), C, Ne) As Variant()
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C) <> Ne Then PushI DywColNe, Dr
Next
End Function

Function DywDist(Dy()) As Variant()
Dim O(), Dr
For Each Dr In Itr(Dy)
    PushNDupDr O, Dr
Next
DywDist = O
End Function

Function DywDup1(Dy(), C&) As Variant()
Dim Dup$(), Dr, O()
Dup = CvSy(AwDup(StrColzDy(Dy, C)))
For Each Dr In Itr(Dy)
    If HasEle(Dup, Dr(C)) Then PushI DywDup1, Dr
Next
End Function

Function DywDupC(Dy(), Coxiy&()) As Variant()
Dim Dup$(), Dr
Dup = AwDup(JnDyCC(Dy, Coxiy))
For Each Dr In Itr(Dy)
    If HasEle(Dup, Jn(AwIxy(Dr, Coxiy), vbFldSep)) Then Push DywDupC, Dr
Next
End Function

Function DywDupzC(Dy(), C&) As Variant()
DywDupzC = AwIxy(Dy, IxyzDup(ColzDy(Dy, C)))
End Function

Function DywEq(Dy(), C&, Eq) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(C) = Eq Then PushI DywEq, Dr
Next
End Function

Function DywEqVy(Dy(), Ixy&(), Vy()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If IsEqAy(AwIxy(Dr, Ixy), Vy) Then PushI DywEqVy, Dr
Next
End Function

Function DywIn(Dy(), C, InVy) As Variant()
Const CSub$ = CMod & "DywIn"
If Not IsArray(InVy) Then Thw CSub, "Given InVy is not an array", "Ty-InVy", TypeName(InVy)
Dim Dr
For Each Dr In Itr(Dy)
    If HasEle(InVy, Dr(C)) Then
        PushI DywIn, Dr
    End If
Next
End Function

Function DywLik(Dy(), C&, Lik) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(C) Like Lik Then PushI DywLik, Dr
Next
End Function

Function DywPfx(Dy(), C&, Pfx, Optional Cmp As VbCompareMethod = vbTextCompare) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
   If HasPfx(Dr(C), Pfx, Cmp) Then PushI DywPfx, Dr
Next
End Function

Function DywSubStr(Dy(), C&, SubStr) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    If HasSubStr(Dr(C), SubStr) Then PushI DywSubStr, Dr
Next
End Function

Function DyzSqColnoy(Sq(), Colnoy%()) As Variant()
'Fm Colnoy : :Colnoy ! selecting which col of @Sq
Dim R&: For R = 1 To UBound(Sq, 1)
    PushI DyzSqColnoy, DrzSqrColnoy(Sq, R, Colnoy)
Next
End Function

Function DyzSq(Sq()) As Variant()
'Fm Colnoy : :Colnoy ! selecting which col of @Sq
Dim R&: For R = 1 To UBound(Sq, 1)
    PushI DyzSq, DrzSqr(Sq, R)
Next
End Function

Function DyzVbl(Vbl$) As Variant()
Dim L: For Each L In Itr(SplitVBar(Vbl))
    PushI DyzVbl, SyzSS(L)
Next
End Function

Function FstDrzDy(Dy(), C, V) As Variant()
Const CSub$ = CMod & "FstDrzDy"
Dim Dr: For Each Dr In Itr(Dy)
    If Dr(C) = V Then FstDrzDy = Dr: Exit Function
Next
Thw CSub, "No first Dr in Dy of Cix eq to V", "Cix V Dy", C, V, JnDy(Dy)
End Function

Function FstRecEqzDy(Dy(), C, Eq, SelIxy&()) As Variant()
Const CSub$ = CMod & "FstRecEqzDy"
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C) = Eq Then FstRecEqzDy = Array(AwIxy(Dr, SelIxy)): Exit Function
Next
Thw CSub, "No first rec in Dy of Col-A eq to Val-B", "Col-A Val-B Dy", C, Eq, JnDy(Dy)
End Function

Function HasColEqzDy(Dy(), C&, Eq) As Boolean
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C) = Eq Then HasColEqzDy = True: Exit Function
Next
End Function

Function HasDr(Dy(), Dr) As Boolean
Dim Idr
For Each Idr In Itr(Dy)
    If IsEqAy(Idr, Dr) Then HasDr = True: Exit Function
Next
End Function

Function HasDrzIxy(Dy(), Dr, Ixy&()) As Boolean
Dim Idr, A()
For Each Idr In Itr(Dy)
    A = AwIxy(Idr, Ixy)
    If IsEqAy(A, Idr) Then HasDrzIxy = True: Exit Function
Next
End Function

Function IsRowBrk(Dy(), R&, BrkColIx&) As Boolean
If Si(Dy) = 0 Then Exit Function
If R& = 0 Then Exit Function
If R = UB(Dy) Then Exit Function
If Dy(R)(BrkColIx) = Dy(R - 1)(BrkColIx) Then Exit Function
IsRowBrk = True
End Function

Function IxyzColnoy(Colnoy) As Long()
If Si(Colnoy) = 0 Then Exit Function
Dim Cno: For Each Cno In Colnoy
    PushI IxyzColnoy, Cno - 1
Next
End Function

Function KeepFstNCol(Dy(), N%) As Variant()
Dim Dr, U%
U = N - 1
For Each Dr In Itr(Dy)
    ReDim Preserve Dr(U)
    PushI KeepFstNCol, Dr
Next
End Function

Function KeepFstNColzDrs(A As Drs, N%) As Drs
KeepFstNColzDrs = Drs(CvSy(FstNEle(A.Fny, N)), KeepFstNCol(A.Dy, N))
End Function

Function NColzDy%(Dy())
Dim O%, Dr
For Each Dr In Itr(Dy)
    O = Max(O, Si(Dr))
Next
NColzDy = O
End Function

Function NRowOfInDyoColEv&(Dy(), C&, Ev)
Dim J&, O&, Dr
For Each Dr In Itr(Dy)
   If Dr(C) = Ev Then O = O + 1
Next
NRowOfInDyoColEv = O
End Function

Function RowIxOptzDyDr&(Dy(), Dr)
Dim N%: N = Si(Dr)
Dim Ix&, D
For Each D In Itr(Dy)
    If IsEqAy(FstNEle(D, N), Dr) Then
        RowIxOptzDyDr = Ix
        Exit Function
    End If
    Ix = Ix + 1
Next
RowIxOptzDyDr = -1
End Function

Function SeqCntDizDy(Dy(), C&) As Dictionary
Set SeqCntDizDy = SeqCntDi(ColzDy(Dy, C&))
End Function

Sub ChkDyEQ(Dy(), B(), Optional Fun$)
If Not IsEqDy(Dy, B) Then Thw Fun, "2 Dy not Eq", "Dy1 Dy2", FmtDy(Dy), FmtDy(B)
End Sub

Private Sub DyJnFldKK__Tst()
Dim Dy(), KKIxy&(), JnFldIx&, Sep$
Sep = " "
Dy = Array(Array(1, 2, 3, 4, "Dy"), Array(1, 2, 3, 6, "B"), Array(1, 2, 2, 8, "C"), Array(1, 2, 2, 12, "DD"), _
Array(2, 3, 1, 1, "x"), Array(12, 3), Array(12, 3, 1, 2, "XX"))
Ept = Array()
KKIxy = Array(0, 1, 2)
JnFldIx = 4
GoSub Tst
Exit Sub
Tst:
    Act = DyJnFldKK(Dy, KKIxy, JnFldIx, Sep)
    BrwDy CvAv(Act)
    StopNE
    Return
End Sub

Private Sub FmtA__Tst()
DmpAy JnDy(SampDy2)
End Sub

Function DyzAyV(Ay, V) As Variant()
Dim I: For Each I In Itr(Ay)
    PushI DyzAyV, Array(I, V)
Next
End Function

Function DyzVAy(V, Ay) As Variant()
Dim I: For Each I In Itr(Ay)
    PushI DyzVAy, Array(V, I)
Next
End Function

Function IsEqCol(Dy(), C1%, C2%) As Boolean ' Is col-@C1 equal col-@C2
Dim M%: M = Max(C1, C2)
Dim Dr: For Each Dr In Itr(Dy)
    If UB(Dr) < M Then Exit Function  ' Dr may have less field than C1,C2
    If Dr(C1) <> Dr(C2) Then Exit Function
Next
IsEqCol = True
End Function
