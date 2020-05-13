Attribute VB_Name = "MxVbDtaIs"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay"
Const CMod$ = CLib & "MxVbDtaIs."
':FunPfx-Am: :Fun-Pfx #Ay-Ele-XXX# ! Given @Ay will return same number of ele after doing some mapping
Public Const StmtBrkPatn$ = "(\.  |\r\n|\r)"
#If Doc Then
'Ay:Cml :Var() #Array# ele can be object
'IAy:Cml :Var() #Item-Ay# its ele cannot be object
':Sy: :String-Array #String-Array#
':SS: :Ln #Spc-Separated# ! It will be separated in :Sy
':IntSeg: :IntAy #Int-Sequence# ! Each next element is always 1 more than previous one
':LngSeg: :LngAy #Int-Sequence# ! Each next element is always 1 more than previous one
#End If

Sub AsgAy(Ay, ParamArray OAp())
Dim OAv(): OAv = OAp
Dim J%: For J = 0 To Min(UB(Ay), UB(OAv))
    OAp(J) = Ay(J)
Next
End Sub

Sub AsgT1AyRstAy(Ly$(), OT1Sy$(), ORestSy$())
OT1Sy = AmT1(Ly)
ORestSy = AmRmvT1(Ly)
End Sub

Sub VcAy(Ay, Optional FnPfx$): LisAy Ay, Vcg(FnPfx): End Sub
Sub BrwAy(Ay, Optional FnPfx$): LisAy Ay, Brwg(FnPfx): End Sub
Sub LisAy(Ay, Oup As OupOpt)
If Oup.OupTy = eDmpOup Then Dmp Ay: Exit Sub
Dim T$: T = TmpFt("BrwAy", Oup.FnPfx)
WrtAy Ay, T
BrwFt T, Oup.OupTy = eVcOup
End Sub


Function DupEleMsgl$(Ay, QMsg$)
Dim Dup: Dup = AwDup(Ay)
If Si(Dup) = 0 Then Exit Function
DupEleMsgl = FmtQQ(QMsg, JnSpc(Dup))
End Function

Function DupAmT1(Ly$(), Optional C As VbCompareMethod = vbTextCompare) As String()
Dim A$(): A = AmT1(Ly)
DupAmT1 = AwDup(A, C)
End Function

Function ChkEmpAy(Ay, Msg$) As String()
If Si(Ay) = 0 Then ChkEmpAy = Sy(Msg)
End Function

Private Sub AyzFlat__Tst()
Dim AyOfAy()
AyOfAy = Array(SyzSS("a b c d"), SyzSS("a b c"))
Ept = SyzSS("a b c d a b c")
GoSub Tst
Exit Sub
Tst:
    Act = AyzFlat(AyOfAy)
    C
    Return
End Sub

Function AyzFlat(AyOfAy())
AyzFlat = AyzAyOfAy(AyOfAy)
End Function

Function ItmCnt%(Ay, M)
If Si(Ay) = 0 Then Exit Function
Dim O%, X
For Each X In Itr(Ay)
    If X = M Then O = O + 1
Next
ItmCnt = O
End Function

Function ResiN(Ay, N&)
'Ret : empty ay of si @N of sam base ele as @Ay @@
ResiN = ResiAy(Ay, N - 1)
End Function

Function IsIn(V, Ay) ' ret True if @V is in @Ay
'Is::Verb
IsIn = HasEle(Ay, V)
End Function

Sub ResiMax(OAy1, OAy2) ' resi the min si of ay to sam si as the other @@
'Resi::Verb resize
Dim U1&, U2&: U1 = UB(OAy1): U2 = UB(OAy2)
Select Case True
Case U1 > U2: ReDim Preserve OAy2(U1)
Case U2 > U1: ReDim Preserve OAy1(U2)
End Select
End Sub
Function NwAy(Ay)
'Nw::Verb New
NwAy = Ay: Erase NwAy
End Function

Function RevAy(Ay) ' rev an @Ay
Dim O: O = Ay
Dim U&: U = UB(O)
Dim J&: For J = 0 To U
    Asg Ay(U - J), O(J)
Next
RevAy = O
End Function

Function RevIAy(IAy) ' rev an @IAy
Dim O: O = IAy
Dim U&: U = UB(O)
Dim J&: For J = 0 To U
    O(J) = IAy(U - J)
Next
RevIAy = O
End Function

Function SeqCntDi(Ay) As Dictionary 'The return dic of key=AyEle pointing to 2-Ele-LngAy with Ele-0 as Seq#(0..) and Ele- as Cnt
Dim S&, O As New Dictionary, L&(), X
For Each X In Itr(Ay)
    If O.Exists(X) Then
        L = O(X)
        L(1) = L(1) + 1
        O(X) = L
    Else
        ReDim L(1)
        L(0) = S
        L(1) = 1
        O.Add X, L
    End If
Next
Set SeqCntDi = O
End Function
Function StrColzSq(Sq(), Optional C = 1) As String()
If Si(Sq) = 0 Then Exit Function
Dim R&: For R = 1 To UBound(Sq, 1)
    PushI StrColzSq, Sq(R, C)
Next
End Function
Function SqhzAp(ParamArray Ap()) As Variant()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
SqhzAp = Sqh(Av)
End Function

Function Sqh(Ay) As Variant()
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To 1, 1 To N)
For Each V In Ay
    J = J + 1
    O(1, J) = V
Next
Sqh = O
End Function

Function Sqv(Ay) As Variant()
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To N, 1 To 1)
For Each V In Ay
    J = J + 1
    O(J, 1) = V
Next
Sqv = O
End Function

Function IndtSy(Sy$(), Optional Indt% = 4) As String()
Dim I, S$
S = Space(Indt)
For Each I In Itr(Sy)
    PushI IndtSy, S & I
Next
End Function

Function WdtzAy%(Ay)
Dim O%, V
For Each V In Itr(Ay)
    O = Max(O, Len(V))
Next
WdtzAy = O
End Function

Function StmtLy(StmtLin) As String()
StmtLy = EnsAyDotSfx(SyzLTrim(Split(StmtLin, ". ")))
End Function
Private Sub AyAsgAy__Tst()
Dim O%, Ay$
AsgAy Array(234, "abc"), O, Ay
Ass O = 234
Ass Ay = "abc"
End Sub

Sub ChkEqAy__Tst()
ChkEqAy Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4)
End Sub

Private Sub MaxEle__Tst()
Dim Ay()
Dim Act
Act = MaxEle(Ay)
Stop
End Sub

Private Sub MinusAy__Tst()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = MinusAy(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
ChkEq Exp, Act
'
Act = MinusAyAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
ChkEq Exp, Act
End Sub

Private Sub SyzAy__Tst()
Dim Act$(): Act = SyzAy(Array(1, 2, 3))
Ass Si(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub AmTrim__Tst()
DmpAy AmTrim(Sy(1, 2, 3, "  a"))
End Sub


Private Sub DupEleMsgl__Tst()
Dim Ay
Ay = Array("1", "1", "2")
Ept = Sy("This item[1] is duplicated")
GoSub Tst
Exit Sub
Tst:
    Act = DupEleMsgl(Ay, "This item[?] is duplicated")
    C
    Return
End Sub

Private Sub HasDupEle__Tst()
Ass HasDupEle(Array(1, 2, 3, 4)) = False
Ass HasDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub

Private Sub AyInsAy__Tst()
Dim Act, Exp, Ay(), B(), At&
Ay = Array(1, 2, 3, 4)
B = Array("X", "Z")
At = 1
Exp = Array(1, "X", "Z", 2, 3, 4)

Act = InsAy(Ay, B, At)
Ass IsEqAy(Act, Exp)
End Sub

Private Sub MinEleus6__Tst()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = MinusAy(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
ChkEqAy Exp, Act
'
Act = MinusAyAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
ChkEqAy Exp, Act
End Sub

Private Sub SyzAy2__Tst()
Dim Act$(): Act = SyzAy(Array(1, 2, 3))
Ass Si(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub AmTrim2__Tst()
DmpAy AmTrim(Sy(1, 2, 3, "  a"))
End Sub

Private Sub KKCMiDy__Tst()
Dim Dy(), Act As KKCntMulItmColDy, KKColIx%(), IxzAy%
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
'Ass Si(Act) = 4
'Ass IsEqAy(Act(0), Array("Ay", 3, 1, 2, 3))
'Ass IsEqAy(Act(1), Array("B", 3, 2, 3, 4))
'Ass IsEqAy(Act(2), Array("C", 0))
'Ass IsEqAy(Act(3), Array("D", 1, "X"))
End Sub

Private Sub AddPfxzSslIn__Tst()
Dim Ssl$, Exp$(), Pfx$
Ssl = "B C D"
Pfx = "A"
Exp = SyzSS("AB AC AD")
GoSub Tst
Exit Sub
Tst:
    Dim Act$()
    Act = AddPfxzSslIn(Pfx, Ssl)
    Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Function AddPfxzSslIn(Pfx$, SsLin) As String()
AddPfxzSslIn = AmAddPfx(SyzSS(SsLin), Pfx)
End Function

Function SpcSepStr$(S)
If S = "" Then SpcSepStr = ".": Exit Function
SpcSepStr = QuoF(EscSqBkt(SlashCrLf(EscBackSlash(S))))
End Function

Function RevSS$(SS)
If SS = "." Then Exit Function
RevSS = UnTidleSpc(UnSlashTab(UnSlashCrLf(SS)))
End Function

Function SslzDr$(Dr)
Dim J&, U&, O$()
U = UB(Dr)
If U < 0 Then Exit Function
ReDim O(U)
For J = 0 To U
    O(J) = SpcSepStr(Dr(J))
Next
SslzDr = JnSpc(O)
End Function

Function ItrzSS(SS)
Asg Itr(SyzSS(SS)), ItrzSS
End Function

Function SrtSS$(SS$) ' return sorted @SS
SrtSS = JnSpc(SrtAy(SyzSS(SS)))
End Function

Function Colnoy(Colny$(), CC) As Integer() ' Column no array of @CC according @Colny.  Colno starts from 1
Dim C: For Each C In Itr(SyzSS(CC))
    Dim Ix&: Ix = EleIx(Colny, C): If Ix = -1 Then Thw CSub, "A Col of @CC not in @ColNy", "C CC ColNy", C, CC, Colny
    PushI Colnoy, Ix + 1
Next
End Function

Function CntzSS%(SS)
CntzSS = Si(SyzSS(SS))
End Function

Function SyzSS(SS) As String() ' split @SS by spc.  replace . as ""
Dim S: For Each S In Itr(SplitSpc(RplDblSpc(Trim(SS))))
    If S = "." Then
        PushI SyzSS, ""
    Else
        PushI SyzSS, S
    End If
Next
End Function

Function IsEqSy(A$(), B$()) As Boolean
If Not IsEqSi(A, B) Then Exit Function
Dim J&, X
For Each X In Itr(A)
    If X <> B(J) Then Exit Function
    J = J + 1
Next
IsEqSy = True
End Function

Function IsEqDr(A, B) As Boolean
Dim X, J&
For Each X In Itr(A)
    If X <> B(J) Then Exit Function
    J = J + 1
Next
IsEqDr = True
End Function

Function IsLEzAy(Ay1, Ay2, IsDesAy() As Boolean) As Boolean
Dim J&: For J = 0 To UB(Ay1)
    If IsDesAy(J) Then
        If Ay1(J) < Ay2(J) Then Exit Function
        If Ay1(J) > Ay2(J) Then IsLEzAy = True: Exit Function
    Else
        If Ay1(J) > Ay2(J) Then Exit Function
        If Ay1(J) < Ay2(J) Then IsLEzAy = True: Exit Function
    End If
Next
IsLEzAy = True
End Function

Function IsGTzAy(Ay1, Ay2) As Boolean
Dim J&: For J = 0 To UB(Ay1)
    If Ay1(J) <= Ay2(J) Then Exit Function
Next
IsGTzAy = True
End Function


Function IsEqAy(A, B) As Boolean
If Not IsArray(A) Then Exit Function
If Not IsArray(B) Then Exit Function
If Not IsEqSi(A, B) Then Exit Function
Dim J&, X
For Each X In Itr(A)
    If Not IsEq(X, B(J)) Then Exit Function
    J = J + 1
Next
IsEqAy = True
End Function

Function AyzAyOfAy(AyOfAy())
If Si(AyOfAy) = 0 Then AyzAyOfAy = EmpAv: Exit Function
AyzAyOfAy = AyOfAy(0)
Dim J&: For J = 1 To UB(AyOfAy)
    PushAy AyzAyOfAy, AyOfAy(J)
Next
End Function

Function SyzLTrim(Ay) As String()
Dim L
For Each L In Itr(Ay)
    PushI SyzLTrim, LTrim(L)
Next
End Function

Function SyzTrim(Ay) As String()
Dim L: For Each L In Itr(Ay)
    PushI SyzTrim, Trim(L)
Next
End Function

Function AddSS(Sy$(), SS$) As String()
AddSS = SyzAp(Sy, SyzSS(SS))
End Function

Function ItrzAwRmvT1(Ay, T1)
Asg Itr(AwRmvT1(Ay, T1)), ItrzAwRmvT1
End Function

Function IxyzAyPatn(Ay, Patn$) As Long()
IxyzAyPatn = IxyzAyRe(Ay, Rx(Patn))
End Function

Function IxyzAyRe(Ay, B As RegExp) As Long()
If Si(Ay) = 0 Then Exit Function
Dim I, O&(), J&
For Each I In Ay
    If B.Test(I) Then Push O, J
    J = J + 1
Next
IxyzAyRe = O
End Function

Function IxyzCC(D As Drs, CC$) As Long()
IxyzCC = IxyzFF(D.Fny, CC)
End Function

Function IxyzFF(Fny$(), FF$) As Long()
IxyzFF = IxyzSubFny(Fny, FnyzFF(FF))
End Function

Function IxyzSubAy(Ay, SubAy, Optional ThwNFnd As Boolean) As Long()
Const CSub$ = CMod & "IxyzSubAy"
Dim E, Ix&
For Each E In SubAy
    Ix = IxzAy(Ay, E)
    If ThwNFnd Then
        If Ix = -1 Then
            Thw CSub, "Ele in SubAy not found in Ay", "Ele SubAy Ay", E, SubAy, Ay
        End If
    End If
    PushI IxyzSubAy, Ix
Next
End Function


Function IxyzSubFny(Fny$(), SubFny$()) As Long()
Dim F: For Each F In Itr(SubFny)
    Dim I&: I = IxzAy(Fny, F)
    If I >= 0 Then PushI IxyzSubFny, I
Next
End Function

Function LikAyzKssAy(KssAy$()) As String()
Dim Kss: For Each Kss In Itr(KssAy)
    PushIAy LikAyzKssAy, Kss
Next
End Function


Private Sub FmtCntDi__Tst()
Dim Ay
GoSub Z
Exit Sub
Z:
    Ay = Array(1, 2, 2, 2, 3, "skldflskdfsdklf" & vbCrLf & "ksdlfj")
    Brw FmtCntDi(CntDi(Ay))

End Sub
