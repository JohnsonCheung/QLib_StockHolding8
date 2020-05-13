Attribute VB_Name = "MxVbStrRxIdr"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbDtaIdr."
Public Const VbKwSS$ = "Function Sub Then If As For To Each End While Wend Loop Do Static Dim Option Explicit Compare Text"
Private Sub NyzStr__Tst()
Dim S$
Dim Text
GoSub Z
'GoSub T0
Exit Sub
Z:
    Dim Lines$: Lines = SrclP
    Dim Ny1$(): Ny1 = NyzStr(Lines)
    Dim Ny2$(): Ny2 = Idry(Lines)
    If Not IsEqAy(Ny1, Ny2) Then Stop
    Return
T0:
    S = "S_S"
    Ept = Sy("S_S")
    GoTo Tst
Tst:
    Act = NyzStr(S)
    C
    Return
End Sub

Private Sub NmSet__Tst()
VcAet SrtAet(NmAet(SrclP))
End Sub

Function NmAet(S) As Dictionary
Set NmAet = Aet(NyzStr(S))
End Function

Function NyzStr(S) As String()
NyzStr = AwNm(SyzSS(RplLf(RplCr(RplPun(S)))))
End Function



Function MthExtny(MthPjDotMdn, PubMthLy$(), PubMthn_To_PjDotModNy As Dictionary) As String()
Dim Cxt$: Cxt = JnSpc(MthCxt(PubMthLy))
Dim Ny$(): Ny = NyzStr(Cxt)
Dim Nm
For Each Nm In Itr(Ny)
    If PubMthn_To_PjDotModNy.Exists(Nm) Then
        Dim PjDotModNy$():
            PjDotModNy = AeEle(PubMthn_To_PjDotModNy(Nm), MthPjDotMdn)
        If HasEle(PjDotModNy, Nm) Then
            PushI MthExtny, Nm
        End If
    End If
Next
End Function

Property Get VbKwAy() As String()
Static X$()
If Si(X) = 0 Then
    X = SyzSS(VbKwSS)
End If
VbKwAy = X
End Property

Property Get VbKwAet() As Dictionary
Set VbKwAet = Aet(VbKwAy)
End Property

'There will be 3 subMch for these patn (()|()): SubMch1 is the outer bkt and SubMch2 and 3 are the inner.  If SubMch2 then 3 will be empty, of SubMch3, 2 will be empty.
Private Sub DiIdrqCnt__Tst()
Dim A As Dictionary: Set A = DiIdrqCnt(JnCrLf(SrczP(CPj)))
Set A = SrtDic(A)
BrwDic A
End Sub
Private Sub IdrStszLines__Tst()
Debug.Print IdrStszLy(SrczP(CPj))
End Sub

Function IdrStszLines$(Lines)
IdrStszLines = IdrStszLy(SplitCrLf(Lines))
End Function

Sub CntIdrP()
Debug.Print IdrStszLy(SrczP(CPj))
End Sub

Function IdrStszLy$(Ly$())
Dim W&, D&, Sy$(), B&, L&, S$
S = JnCrLf(Ly)
Sy = Idry(S)
W = Si(Sy)
D = Si(AwDis(Sy))
L = Si(Ly)
B = Len(S)
IdrStszLy = IdrSts(B, L, W, D)
End Function

Function IdrSts$(B&, L&, W&, D&)
Dim BB As String * 9: RSet BB = B
Dim LL As String * 9: RSet LL = L
Dim WW As String * 9: RSet WW = W
Dim DD As String * 9: RSet DD = D
IdrSts = FmtQQ("Len            : ?|Lines          : ?|Words          : ?|Distinct Words : ?", BB, LL, WW, DD)
End Function

Function NIdr&(S)
NIdr = Si(Idry(S))
End Function

Function NDistIdr&(S)
NDistIdr = Si(AwDis(Idry(S)))
End Function

Function DiIdrqCnt(S) As Dictionary
Set DiIdrqCnt = CntDi(Idry(S))
End Function

Function IdrAet(S) As Dictionary
Set IdrAet = Aet(Idry(S))
End Function

Function CvMch(A) As Match:        Set CvMch = A:  End Function
Function CvSMchs(A) As SubMatches: Set CvSMchs = A: End Function

Function FstIdrAetP() As Dictionary
Set FstIdrAetP = New Dictionary
Dim L: For Each L In Itr(RmvVrmk(SrczP(CPj)))
    PushEle FstIdrAetP, FstIdr(L)
Next
End Function
'--
Private Sub FstIdr__Tst()
Dim S$
GoSub T1
Exit Sub
T1:
    S = "00A cB"
    Ept = "cB"
    GoTo Tst
Tst:
    Act = FstIdr(S)
    C
    Return
End Sub

Function IdrRx() As RegExp
Const P$ = "(^[A-Z]\w*)|[ .\(]([A-Z]\w*)" ' Rx should use ignorecas.  2 cases: a word is begin of a line or fst chr is one of these char [ .(]..
Dim X As RegExp: If IsNothing(X) Then Set X = Rx(P, MultiLine:=True, IgnoreCase:=True)
Set IdrRx = X
End Function
Function FstIdr$(S) ' first Idr.  Word is using regexp ZZIdrPatn1
FstIdr = MchszR(S, IdrRx)
End Function

Function NNmChrRx() As RegExp ' Non name char regexp
Dim O As New RegExp
Set O = Rx("\W")
End Function

Function RplNNmChr$(S) ' replace non name char to space
'NNmChr:Cml #non-nm-chr#
RplNNmChr = NNmChrRx.Replace(S, " ")
End Function

Function AmRplNNmChr(Ay) As String()
Dim L: For Each L In Itr(Ay)
    PushI AmRplNNmChr, RplNNmChr(L)
Next
End Function

Private Sub AmRplNNmChr__Tst()
Brw AmRplNNmChr(SrczP(CPj))
End Sub

'--
Private Sub IdryzS__Tst()

End Sub

Function IdryzS(Src$()) As String()
Dim L: For Each L In Itr(RmvVrmkAndVstr(Src))
    PushI IdryzS, Idry(L)
Next
End Function

'--
Private Sub Idry__Tst()
Dim S$
'GoSub T1
GoSub ZZ
Exit Sub
ZZ:
    VcAy Idry(SrclP)
    Return
T1:
    S = "Function 0AA B"
    Ept = Sy("Function", "B")
    GoTo Tst
Tst:
    Act = Idry(S)
    C
    Return
End Sub

Function Idry(S) As String() ' Identifier array of @Lines
Dim M As Match: For Each M In IdrRx.Execute(S)
    PushI Idry, ZZIdrzMch(M)
Next
End Function

Function HasIdr(S, Idr) As Boolean
HasIdr = HasEle(Idry(Sy(S)), Idr)
End Function

Function Idrss$(S)
Idrss = JnSpc(Idry(S))
End Function

Function IdrssAy(Sy$()) As String()
Dim S: For Each S In Itr(Sy)
    PushI IdrssAy, Idrss(S)
Next
End Function


Function IdrLblLinPos$(IdrPos%(), OFmNo&)
Dim O$(), A$, B$, W%, J%
If Si(IdrPos) = 0 Then Exit Function
PushNB O, Space(IdrPos(0) - 1)
For J = 0 To UB(IdrPos) - 1
    A = OFmNo
    W = IdrPos(J + 1) - IdrPos(J)
    If W > Len(A) Then
        A = AliL(A, W)
        If Len(A) <> W Then Stop
    Else
        A = Space(W)
    End If
    PushI O, A
    OFmNo = OFmNo + 1
Next
A = OFmNo
PushI O, A
IdrLblLinPos = Jn(O)
End Function
Function IdrLblLin(Ln, OFmNo&)
IdrLblLin = IdrLblLinPos(IdrPosy(Ln), OFmNo)
End Function

Function IdrPosy(Ln) As Integer()
Dim J%, LasIsSpc As Boolean, CurIsSpc As Boolean
LasIsSpc = True
For J = 1 To Len(Ln)
    CurIsSpc = Mid(Ln, J, 1) = " "
    Select Case True
    Case CurIsSpc And LasIsSpc
    Case CurIsSpc:          LasIsSpc = True
    Case LasIsSpc:          PushI IdrPosy, J
                            LasIsSpc = False
    Case Else
    End Select
Next
End Function
Function IdrLblLinPairLno(Ln, Lno, LnoWdt, OFmNo&) As String()
Dim O$(): O = IdrLblLinPair(Ln, OFmNo)
O(0) = Space(LnoWdt) & " : " & O(0)
'O(1) = AliL(Lno, LnoWdt) & " : " & O(1)
IdrLblLinPairLno = O
End Function
Function IdrLblLinPair(Ln, OFmNo&) As String()
PushI IdrLblLinPair, IdrLblLin(Ln, OFmNo)
PushI IdrLblLinPair, Ln
End Function
Function IdrLblLy(Ly$(), OFmNo&) As String()
Dim J&, LnoWdt%, A$
A = UB(Ly)
LnoWdt = Len(A)
For J = 1 To UB(Ly)
    PushIAy IdrLblLy, IdrLblLinPairLno(Ly(J), J, LnoWdt, OFmNo)
Next
End Function


Private Sub IdrLblLin__Tst()
Dim Ln, FmNo&
GoSub T0
Exit Sub
T0:
    FmNo = 2
    '               10        20        30        40        50        60
    '      123456789 123456789 123456789 123456789 123456789 123456789 123456789
    Ln = "Lbl01 Lbl02 Lbl03    Lbl04 Lbl05 Lbl06 Lbl07 Lbl08 Lbl09 Lbl10"
    Ept = "2     3     4        5     6     7     8     9     10    11"
    GoTo Tst
Tst:
    Act = IdrLblLin(Ln, FmNo)
    C
    Return
End Sub
Private Sub IdrPosy__Tst()
Dim Ln
GoSub T0
Exit Sub
T0:
    '               10        20        30        40        50        60
    '      123456789 123456789 123456789 123456789 123456789 123456789 123456789
    Ln = "Lbl01 Lbl02 Lbl03    Lbl04 Lbl05 Lbl06 Lbl07 Lbl08 Lbl09 Lbl10"
    Ept = IntAy(1, 7, 13, 22, 28, 34, 40, 46, 52, 58)
    GoTo Tst
Tst:
    Act = IdrPosy(Ln)
    C
    Return
End Sub

Private Sub IdrLblLy__Tst()
Dim Fm&: Fm = 1
Brw IdrLblLy(SrczP(CPj), Fm)
End Sub

Function IdryP() As String(): IdryP = IdryzP(CPj): End Function

Function IdryzP(P As VBProject) As String()
Dim L: For Each L In SrczP(P)
    PushIAy IdryzP, Idry(L)
Next
End Function
'== ZZ
Private Function ZZIdrzMch$(M As Match)
Dim S As ISubMatches: Set S = M.SubMatches
If S.Count <> 2 Then Imposs CSub, "The ZZIdrPatn should be ()|() so that it will gives 2 subMatch, but now the submatch count=[" & S.Count & "]"
If IsEmpty(S(1)) Then
    ZZIdrzMch = S(0) ' fst-SubMch match means the idr started with begin of a line
Else
    ZZIdrzMch = S(1)  ' snd-SubMch match means the idr started with 1 of spc of (
End If
End Function
