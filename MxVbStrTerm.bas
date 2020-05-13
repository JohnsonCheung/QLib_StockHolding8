Attribute VB_Name = "MxVbStrTerm"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Dta"
Const CMod$ = CLib & "MxVbStrTerm."
Enum eSepOpt: eExlSep: eInlSep: End Enum
Sub AsgTT(Ln, O1, O2):                       AsgAy T2Rst(Ln), O1, O2:               End Sub
Sub AsgTRst(Ln, OT1, ORst):                  AsgAy SyzTRst(Ln), OT1, ORst:          End Sub
Sub AsgTTRst(Ln, OT1, OT2, ORst$):           AsgAy T2Rst(Ln), OT1, OT2, ORst:       End Sub
Sub Asg3TRst(Ln, OT1, OT2, OT3, ORst$):      AsgAy T3Rst(Ln), OT1, OT2, OT3, ORst:  End Sub
Sub Asg4T(Ln, O1$, O2$, O3$, O4$):           AsgAy Fst4Term(Ln), O1, O2, O3, O4:    End Sub
Sub Asg4TRst(Ln, O1$, O2$, O3$, O4$, ORst$): AsgAy T4Rst(Ln), O1, O2, O3, O4, ORst: End Sub

Sub AsgAmT1RstAy(Ly$(), OAmT1$(), ORstAy$())
Erase OAmT1, ORstAy
Dim L: For Each L In Itr(Ly)
    PushI OAmT1, T1(L)
    PushI ORstAy, RmvT1(L)
Next
End Sub

Sub AsgT1Sy(LinOf_T1_SS, OT1, Osy$())
Dim Rst$
AsgTRst LinOf_T1_SS, OT1, Rst
Osy = SyzSS(Rst)
End Sub

Function Fst2Term(Ln) As String(): Fst2Term = FstNTerm(Ln, 2): End Function
Function Fst3Term(Ln) As String(): Fst3Term = FstNTerm(Ln, 3): End Function
Function Fst4Term(Ln) As String(): Fst4Term = FstNTerm(Ln, 4): End Function
Function FstNTerm(Ln, N%) As String()
Dim J%, L$
L = Ln
For J = 1 To N
    PushI FstNTerm, ShfTerm(L)
Next
End Function

Function SyzTRst(Ln) As String()
SyzTRst = NTermRst(Ln, 1)
End Function

Private Sub NTermRst__Tst()
Dim Ln
Ln = "  [ksldfj ]":  Ept = "ksldfj ": GoSub Tst
Ln = "  [ ksldfj ]": Ept = " ksldf ": GoSub Tst
Ln = "  [ksldfj]":  Ept = "ksldf": GoSub Tst
Exit Sub
Tst:
    Act = T1(Ln)
    C
    Return
End Sub
Function T2Rst(Ln) As String(): T2Rst = NTermRst(Ln, 2): End Function
Function T3Rst(Ln) As String(): T3Rst = NTermRst(Ln, 3): End Function
Function T4Rst(Ln) As String(): T4Rst = NTermRst(Ln, 4): End Function
Function NTermRst(Ln, N%) As String() '#NTerm-and-Rest# with @N+1 ele from @L
Dim L$: L = Ln
Dim J%: For J = 1 To N
    PushI NTermRst, ShfTerm(L)
Next
PushI NTermRst, L
End Function

'**TermN
Private Sub TermN__Tst()
Dim N%, A$
N = 1: A = "a b c": Ept = "a": GoSub Tst
N = 2: A = "a b c": Ept = "b": GoSub Tst
N = 3: A = "a b c": Ept = "c": GoSub Tst
Exit Sub
Tst:
    Act = TermN(A, N)
    C
    Return
End Sub
Function FstTerm$(S): FstTerm = T1(S):            End Function
Function T1$(S):           T1 = ShfTerm(CStr(S)): End Function
Function T2zS$(S):       T2zS = T2(S):            End Function
Function T2$(S):           T2 = TermN(S, 2):      End Function
Function T3$(S):           T3 = TermN(S, 3):      End Function
Function TermN$(S, N%)
Dim L$, J%
L = LTrim(S)
For J = 1 To N - 1
    L = RmvT1(L)
Next
TermN = T1(L)
End Function


':Term: :S ! No-spc-str or Sq-quoted-str
':NN: :SS ! #spc-sep-name#
Function RmvT1XzA$(Ln, Termy$())
Dim T$: T = T1(Ln)
If HasEle(Termy, T) Then
    RmvT1XzA = RmvT1x(Ln, T)
Else
    RmvT1XzA = Ln
End If
End Function
Function RmvT1x$(Ln, T1x$)
Dim T$: T = T1(Ln)
If T = T1x Then
    RmvT1x = RmvT1(Ln)
Else
    RmvT1x = LTrim(Ln)
End If
End Function
Function RplT1$(L, T1$, By$)
If HasT1(L, T1) Then
    RplT1 = By & Mid(L, Len(T1) + 1)
Else
    RplT1 = L
End If
End Function

'**Termy
Function TermAet(S) As Dictionary: Set TermAet = Aet(Termy(S)):                End Function
Function TermItr(Tml$):                          Asg Itr(Termy(Tml)), TermItr: End Function
Function ItrzTml(Tml$):                          Asg Itr(Termy(Tml)), ItrzTml: End Function
Function Termy(Termln) As String()
Dim L$, J%: L = Termln
While L <> ""
    LoopTooMuch "Termy", J
    PushNB Termy, ShfTerm(L)
Wend
End Function
Function QuoTermy(Termy) As String()
Dim I: For Each I In Itr(Termy)
    PushI QuoTermy, QuoTerm(I)
Next
End Function

'**Tml
Function JnTerm$(Termy):  JnTerm = JnSpc(QuoTerm(AwNB(Termy))):           End Function
Function Tml$(Termy):        Tml = Termln(Termy):                         End Function
Function Termln$(Termy):  Termln = JnSpc(QuoTermy(Termy)):                End Function  ' #Term-Lin# Fmt is space separated.  If term has space, use [].
Function QuoTerm$(Term): QuoTerm = IIf(HasSpc(Term), QuoSq(Term), Term): End Function
Function TmlzAp$(ParamArray TermAp())
Dim Av(): Av = TermAp: TmlzAp = Tml(Av)
End Function

'**ShfTerm
Function NmAftTerm$(Ln, Term$)
Dim L$: L = Ln
If Not ShfTermX(L, Term) Then Exit Function
NmAftTerm = TakNm(L)
End Function
Function ShfTerm$(OLin)
Dim S$: S = LTrim(OLin)
If S = "" Then Exit Function
If FstChr(S) <> "[" Then
    Dim P%: P = InStr(S, " ")
    If P = 0 Then
        ShfTerm = S
        OLin = ""
        Exit Function
    End If
    ShfTerm = Left(S, P - 1)
    OLin = LTrim(Mid(S, P + 1))
    Exit Function
End If
P = InStr(S, "]")
If P = 0 Then Raise "ShfTerm: Invalid OLin=[" & OLin & "]"
ShfTerm = Mid(S, 2, P - 2)
OLin = LTrim(Mid(S, P + 1))
End Function
Function ShfTermX(OLin$, TermX$) As Boolean
If T1(OLin) <> TermX Then Exit Function
ShfTermX = True
OLin = RmvT1(OLin)
End Function

Function ShfTermDot$(OLin$)
With Brk2Dot(OLin, NoTrim:=True)
    ShfTermDot = .S1
    OLin = .S2
End With
End Function


Private Sub ShfBef__Tst()
Dim L$, Sep$, EptL$, Opt As eSepOpt
GoSub T1a
GoSub T1b
Exit Sub
T1a:
    L = "aaa == bbb"
    Sep = "=="
    Opt = eInlSep
    Ept = "aaa =="
    EptL = " bbb"
    GoTo Tst
T1b:
    L = "aaa == bbb"
    Sep = "=="
    Opt = eExlSep
    Ept = "aaa "
    EptL = "== bbb"
    GoTo Tst
T0:
    L = "aaa.bbb"
    Sep = "."
    Ept = "aaa"
    EptL = ".bbb"
    GoTo Tst
Tst:
    Act = ShfBef(L, Sep, Opt)
    C
    If L <> EptL Then Stop
    Return
End Sub
Function ShfBefEq$(OLn$, Optional Opt As eSepOpt): ShfBefEq = ShfBef(OLn, "=", Opt): End Function
Function ShfBef$(OLn$, Bef, Optional Opt As eSepOpt, Optional NoTrim As Boolean)
Dim P&: P = InStr(OLn, Bef): If P = 0 Then Thw CSub, "Bef not found in OLn", "Bef OLn", Bef, OLn
If Opt = eExlSep Then
    ShfBef = Left(OLn, P - 1)
    OLn = Mid(OLn, P)
Else
    Dim L%: L = Len(Bef)
    ShfBef = Left(OLn, P - 1 + L)
    OLn = Mid(OLn, P + L)
End If
If Not NoTrim Then
    ShfBef = Trim(ShfBef)
    OLn = Trim(OLn)
End If
End Function
Function ShfBefOpt$(OLn$, Bef)
Dim P&: P = InStr(OLn, Bef): If P = 0 Then Exit Function
ShfBefOpt = Left(OLn, P - 1)
OLn = Mid(OLn, P)
End Function
Function ShfBefOrAll$(OLin$, Sep$, Optional NoTrim As Boolean)
Dim P%: P = InStr(OLin, Sep)
If P = 0 Then
    If NoTrim Then
        ShfBefOrAll = OLin
    Else
        ShfBefOrAll = Trim(OLin)
    End If
    OLin = ""
    Exit Function
End If
ShfBefOrAll = Bef(OLin, Sep, NoTrim)
OLin = Aft(OLin, Sep, NoTrim)
End Function

Function FndT1$(Itm, TkssAy$())
Dim L$, I, Kss$, T1$
For Each I In TkssAy
    L = I
    AsgTRst L, T1, Kss
    If HitKss(Itm, Kss) Then FndT1 = T1: Exit Function
Next
End Function

Function Has2T(S, T1, T2) As Boolean
Dim L$: L = S
If ShfTerm(L) <> T1 Then Exit Function
If ShfTerm(L) <> T2 Then Exit Function
Has2T = True
End Function

Function Has3T(S, T1, T2, T3) As Boolean
Dim L$: L = S
If ShfTerm(L) <> T1 Then Exit Function
If ShfTerm(L) <> T2 Then Exit Function
If ShfTerm(L) <> T3 Then Exit Function
Has3T = True
End Function

Function HasT1(S, T1) As Boolean
HasT1 = FstTerm(S) = T1
End Function

Function HasT2(Ln, T2) As Boolean
HasT2 = T2zS(Ln) = T2
End Function
Function TRstSi&(A() As TRst)
On Error Resume Next
TRstSi = UBound(A) + 1
End Function
Function TRstUB&(A() As TRst)
TRstUB = TRstSi(A) - 1
End Function
Function T1Ay(A() As TRst) As String()
Dim J&: For J = 0 To TRstUB(A)
    PushS T1Ay, A(J).T
Next
End Function
Function RstAy(A() As TRst) As String()
Dim J&: For J = 0 To TRstUB(A)
    PushS RstAy, A(J).Rst
Next
End Function

Function TRstAy(Ly$()) As TRst()
Dim L: For Each L In Itr(Ly)
    PushTRst TRstAy, TRstzLn(L)
Next
End Function

Function TRst(T, Rst) As TRst
With TRst
    .T = T
    .Rst = Rst
End With
End Function

Function TRstzLn(Ln) As TRst
Dim L$: L = LTrim(Ln)
TRstzLn = TRst(ShfTerm(L), L)
End Function

Private Sub ShfTerm__Tst()
Dim OLin$, EptOLin$, O$
GoSub Z
Exit Sub
Z1:
    O = " AA BB "
    Ept = "AA"
    EptOLin = "BB "
    GoTo Tst
Z:
    OLin = " sdlfkj sdlkf"
    Ept = "sdlfkj"
    EptOLin = "sdlkf"
    GoSub Tst
    OLin = "   [ kdjf ] sdlkfj1"
    Ept = " kdjf "
    EptOLin = "sdlkfj1"
    GoTo Tst
Tst:
    Act = ShfTerm(OLin)
    If Ept <> Act Then Stop
    If EptOLin <> OLin Then Stop
    Return
End Sub

Sub AsgTml(Tml$, ParamArray OAp())
Dim A$(): A = Termy(Tml)
Dim Av(): Av = OAp
Dim U1%, U2%: U1 = UB(A): U2 = UB(Av)
Dim J%
For J = 0 To U2: OAp(J) = Empty: Next
For J = 0 To Min(U1, U2)
    OAp(J) = A(J)
Next
End Sub

Function FnyzFF(FF$) As String(): FnyzFF = Termy(FF): End Function
