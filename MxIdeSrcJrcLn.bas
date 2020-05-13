Attribute VB_Name = "MxIdeSrcJrcLn"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Jrc"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcJrcLn."
Private Sub Jrc__Tst()
BrwAy JrcP("\sDim J%")
End Sub

Function JrcP(LnPatn$) As String()
JrcP = Jrc(CPj, LnPatn)
End Function

Function Jrc(Pj As VBProject, LnPatn$) As String() ' return Sy of Jrc, which is a srcln will jump to particular src line
Dim P As PjSrc: P = PjSrczP(Pj)
Dim R As RegExp: Set R = Rx(LnPatn)
Dim O$()
Dim J%: For J = 0 To MdSrcUB(P.Md)
    PushIAy O, JrczM(P.Md(J), R)
Next
Jrc = AliLyz1T(O)
End Function
Private Function JrczM(M As MdSrc, Rx As RegExp) As String()
Dim Src$(): Src = M.Src
Dim J&: For J = 0 To UB(Src)
    Dim L$: L = Src(J)
    Dim P As C12: P = C12zRx(L, Rx)
    If Not IsEmpC12(P) Then PushI JrczM, JrclnzP(M.Mdn, J + 1, P, L)
Next
End Function
Function JrclnzP$(Mdn$, Lno&, P As C12, Ln$)
JrclnzP = Jrcln(Mdn, Lno, P.C1, P.C2, Ln)
End Function

Function Jrcln$(Mdn$, Lno&, C1%, C2%, Ln)
Jrcln = FmtQQ("Jmp""?:?:?:?"" '?", Mdn, Lno, C1, C2, Ln)
End Function

Function JrcyzIdr(Idr$) As String()
Stop '
'JSrczIdr = JSrczPred(PredHasIdr(Idr))
End Function

Function JrcyzPfx(LinPfx$) As String()
Stop '
'JSrczPfx = JSrczPred(PredHasPfx(LinPfx))
End Function

Function JrcyzPatn(Patn$, Optional And1$, Optional And2$) As String()
JrcyzPatn = JrcyzRx(Rx(Patn))
End Function

Function JrcyzRx(A As RegExp) As String()
Dim C As VBComponent: For Each C In CPj.VBComponents
    'PushIAy JrcyzRx, JrczRx(C.Name, Src(C.CodeModule), A)
Next
End Function

Function Jrcy(Mdn$, Src$(), A As RegExp) As String()
Dim J&: For J = 0 To UB(Src)
    Dim M As C12: ' M = LinC12(Src(J), A)
    If Not IsEmpC12(M) Then
'       PushI Jrcy, Jrcln(Mdn, J + 1, N.C1, M.C2, Src(J))
    End If
Next
End Function

Private Sub LinC12__Tst()
Dim A As C12: A = LinC12("aAAa", Rx("AA"))
Stop
End Sub

Function LinC12(Ln, A As RegExp) As C12
Dim C As MatchCollection: Set C = Mch(Ln, A): If C.Count = 0 Then Exit Function
Dim M As Match: Set M = C(0)
Dim O As C12
O.C1 = M.FirstIndex + 1
O.C2 = O.C1 + M.Length
LinC12 = O
End Function

Function LinC12zRx(Ln, A As RegExp) As C12
':LinC12: :C12 ! #Ln-C12# Highlighting a portion of a Ln by this C12
'If HitRx(Ln, A) Then
'HitAy
End Function
