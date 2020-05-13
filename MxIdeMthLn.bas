Attribute VB_Name = "MxIdeMthLn"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthln."
'--MthELn
Function MthEixy(Src$(), Bixy&()) As Long()
Dim Ix: For Each Ix In Itr(Bixy)
    PushI MthEixy, MthEix(Src, Ix)
Next
End Function
Function SrcIx&(Src$(), SrclnPfx$, Optional FmIx = 0)
If FmIx < 0 Then SrcIx = -1: Exit Function
Dim O&: For O = FmIx To UB(Src)
   If HasPfx(Src(O), SrclnPfx) Then SrcIx = O: Exit Function
Next
SrcIx = -1
End Function
Function MthEix&(Src$(), Ix)
If Ix < 0 Then MthEix = -1: Exit Function
Dim K$: K = MthKdzL(Src(Ix)): If K = "" Then Thw CSub, "Given Ix in Src is not a method line", "Ix Src(Ix)", Ix, Src(Ix)
Dim Endln$: Endln = "End " & K
If HasSubStr(Src(Ix), Endln) Then MthEix = Ix: Exit Function
MthEix = SrcIx(Src, Endln, Ix)
ThwTrue MthEix < 0, CSub, "Cannot find MthEndLn", "MthEndLin Bix Src", Endln, Ix, Src
End Function

Function MthEno&(M As CodeModule, Mthlno&)
Dim MLn$: MLn = M.Lines(Mthlno, 1)
Dim ELn$: ELn = MthELn(MLn)
Dim O&: For O = Mthlno + 1 To M.CountOfLines
    If HasPfx(M.Lines(O, 1), ELn) Then MthEno = O: Exit Function
Next
End Function

Function MthELn$(Mthln)
Const CSub$ = CMod & "MthELin"
Dim K$: K = MthKdzL(Mthln)
If K = "" Then Thw CSub, "Invalid Mthln", "Mthln", Mthln
MthELn = "End " & K
End Function

'--Mthlny
Function Mthlny(Src$()) As String()
Mthlny = MthlnyzS(Src)
End Function

Function MthlnyM() As String()
MthlnyM = MthlnyzM(CMd)
End Function

Function MthlnyP() As String()
MthlnyP = MthlnyzP(CPj)
End Function

Function MthlnyzP(P As VBProject) As String()
MthlnyzP = MthlnyzS(SrczP(P))
End Function

Function MthlnyzM(M As CodeModule) As String()
MthlnyzM = MthlnyzS(SrczM(M))
End Function

Function MthlnyzS(Src$()) As String()
Dim Ix: For Each Ix In Itr(Mthixy(Src))
    PushI MthlnyzS, MthlnzSI(Src, Ix)
Next
End Function

'-- Mthln
Function MthlnzSI$(Src$(), Ix)
If Not IsMthln(Src(Ix)) Then Exit Function
MthlnzSI = BefColonOrAll(Contln(Src, Ix))
End Function

'--
Private Sub MthlnyzN__Tst()
Dim Mthn$, Src$()
GoSub T1
GoSub T2
Exit Sub
T1:
    Src = SrcM
    Mthn = "MthlnyzN"
    Ept = Sy("Function MthlnyzN(Src$(), Mthn) As String()")
    GoTo Tst
T2:
    Src = SrcM
    Mthn = "AA"
    Ept = Sy("Private Property Get AA$(A, B)", "Private Property Let AA(A, B, V$): End Property")
Tst:
    Act = MthlnyzSNT(Src, Mthn)
    C
    Return
End Sub
Private Property Get AA$(A, _
B)
End Property
Private Property Let AA(A, B, V$): End Property

Function MthlnyzSNT(Src$(), Mthn, Optional ShtMthTy$) As String()
Dim Ix: For Each Ix In Itr(MthixyzSNT(Src, Mthn, ShtMthTy))
    PushI MthlnyzSNT, Contln(Src, Ix)
Next
End Function
Function CMthln$()
Dim M As CodeModule: Set M = CMd
CMthln = ContlnzM(M, MthlnozCLno(M, CLnoM))
End Function

Function MthlnyV() As String()
MthlnyV = MthlnyzV(CVbe)
End Function

Function MthlnyzV(V As Vbe) As String()
Dim P As VBProject: For Each P In V.VBProjects
    PushIAy MthlnyzV, MthlnyzP(P)
Next
End Function

Sub VcMthlAyP()
Vc FmtLinesy(MthlAyP)
End Sub

Function MthlAyP() As String()
MthlAyP = MthlAyzP(CPj)
End Function

Function MthlAyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy MthlAyzP, MthlAyzM(C.CodeModule)
Next
End Function

Function MthlAyzM(M As CodeModule) As String()
MthlAyzM = MthlAyzS(Src(M))
End Function

Function MthlAyzS(Src$()) As String()
Dim Ix: For Each Ix In Itr(Mthixy(Src))
    PushI MthlAyzS, Mthl(Src, Ix)
Next
End Function

Function MdzMthn(P As VBProject, Mthn, Optional ShtMthTy$) As CodeModule
Const CSub$ = CMod & "MdzMthn"
Dim C As VBComponent
For Each C In P.VBComponents
    If HasMthnzMNT(C.CodeModule, Mthn, ShtMthTy) Then
        Set MdzMthn = C.CodeModule
        Exit Function
    End If
Next
Thw CSub, "Mthn not fnd in any codemodule of given pj", "Pj Mthn", "P.Name,Mthn"
End Function

Function MthlnPzNT$(Mthn, Optional ShtMthTy$)
MthlnPzNT = MthlnzPNT(CPj, Mthn, ShtMthTy)
End Function

Function MthlnzPNT$(P As VBProject, Mthn, Optional ShtMthTy$)
Dim C As VBComponent: For Each C In P.VBComponents
    Dim O$: O = MthlnzMNT(C.CodeModule, Mthn, ShtMthTy$)
    If O <> "" Then MthlnzPNT = O: Exit Function
Next
End Function

Function MthlnzMNT$(M As CodeModule, Mthn, Optional ShtMthTy$)
Dim S$(): S = Src(M)
Dim Ix&: Ix = MthixzN(S, Mthn, ShtMthTy)
MthlnzMNT = Contln(S, Ix)
End Function

Function MthlnzSNT$(Src$(), Mthn, Optional ShtMthTy$)
MthlnzSNT = Contln(Src, MthixzN(Src, Mthn, ShtMthTy))
End Function

Function PubMthlny(Src$()) As String()
Dim L, Ix&: For Each L In Itr(Src)
    If IsPubMthln(L) Then PushI PubMthlny, Contln(Src, Ix)
    Ix = Ix + 1
Next
End Function

Function PubMthlnItr(Src$())
Asg Itr(PubMthlny(Src)), PubMthlnItr
End Function

Function NMthln%(M As CodeModule, Mthlno&)
Const CSub$ = CMod & "NMthln"
Dim K$, J&, N&, E$
K = MthKd(M.Lines(Mthlno, 1))
If K = "" Then Thw CSub, "Given Mthlno is not a Mthln", "Md Mthlno Mthln", Mdn(M), Mthlno, M.Lines(Mthlno, 1)
E = "End " & K
For J = Mthlno To M.CountOfLines
    N = N + 1
    If M.Lines(J, 1) = E Then NMthln = N: Exit Function
Next
Imposs CSub
End Function

Function TyChr$(Ln)
'TyChr:: :Chr TyChr of a :Mthln
If IsMthln(Ln) Then TyChr = TakTyChr(DltMthn3(Ln))
End Function

Function MthChr$(Ln)
Dim A$: A = RmvMdy(Ln)
If ShfMthTy(A) = "" Then Exit Function
MthChr = TakTyChr(RmvNm(A))
End Function
