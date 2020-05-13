Attribute VB_Name = "MxVbStrAftBefTak"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str.Op"
Const CMod$ = CLib & "MxVbStrAftBefTak."

Function BefDotRev$(S)
BefDotRev = BefRev(S, ".")
End Function

Function BefDot$(S)
BefDot = Bef(S, ".")
End Function

Function BefDotOrAll$(S)
BefDotOrAll = BefOrAll(S, ".")
End Function

Function BefComma(S, Optional NoTrim As Boolean)
BefComma = Aft(S, vbComma, NoTrim)
End Function

Function BefCommaOrAll(S, Optional NoTrim As Boolean)
BefCommaOrAll = BefOrAll(S, vbComma, NoTrim)
End Function

Function AftCommaOrAll(S, Optional NoTrim As Boolean)
AftCommaOrAll = AftOrAll(S, vbComma, NoTrim)
End Function

Function AftComma(S, Optional NoTrim As Boolean)
AftComma = Aft(S, vbComma, NoTrim)
End Function

Function AftSngQ$(S, Optional NoTrim As Boolean)
AftSngQ = Aft(S, vbSngQ, NoTrim)
End Function

Function Aft$(S, Sep$, Optional NoTrim As Boolean)
Aft = Brk1(S, Sep, NoTrim).S2
End Function

Function AftMust$(S, Sep$, Optional NoTrim As Boolean)
AftMust = Brk(S, Sep, NoTrim).S2
End Function

Function AftColonOrAll$(S)
AftColonOrAll = AftOrAll(S, ":", NoTrim:=True)
End Function

Function BefColonOrAll$(S)
BefColonOrAll = BefOrAll(S, ":", NoTrim:=True)
End Function

Function AftAt$(S, At&, Sep$)
If At = 0 Then Exit Function
AftAt = Mid(S, At + Len(Sep))
End Function

Function AftDotOrAll$(S)
AftDotOrAll = AftOrAll(S, ".")
End Function

Function AftDotOrAllRev$(S)
AftDotOrAllRev = AftOrAllRev(S, ".")
End Function

Function AftDot$(S)
AftDot = Aft(S, ".")
End Function

Function AftOrAll$(S, Sep$, Optional NoTrim As Boolean)
AftOrAll = Brk2(S, Sep, NoTrim).S2
End Function

Function AftOrAllRev$(S, Sep$)
AftOrAllRev = StrDft(AftRev(S, Sep), Sep)
End Function

Function AftRev$(S, Sep$, Optional NoTrim As Boolean)
AftRev = Brk1Rev(S, Sep, NoTrim).S2
End Function

Function BefRev$(S, Sep$, Optional NoTrim As Boolean)
BefRev = Brk1Rev(S, Sep, NoTrim).S1
End Function

Function BefSpc$(S)
BefSpc = Bef(S, " ")
End Function

Function AftSpc$(S, Optional NoTrim As Boolean)
AftSpc = Aft(S, " ", NoTrim)
End Function
Function BefSpcOrAll$(S)
BefSpcOrAll = BefOrAll(S, " ")
End Function
Function BefzSy(Sy$(), Sep$, Optional NoTrim As Boolean) As String()
Dim I
For Each I In Itr(Sy)
    PushI BefzSy, Bef(I, Sep, NoTrim)
Next
End Function

Function BefLowDash$(S)
BefLowDash = Brk2(S, "_", NoTrim:=True).S1
End Function

Function Bet$(S, P1%, P2%)
Bet = Mid(S, P1 + 1, P2 - P1 - 1)
End Function

Function BetP12$(S, P As C12)
If IsEmpC12(P) Then Exit Function
BetP12 = Mid(S, P.C1 + 1, P.C2 - P.C1 - 1)
End Function

Function Bef$(S, Sep$, Optional NoTrim As Boolean)
Bef = Brk2(S, Sep, NoTrim).S1
End Function

Function RmvBef$(S, Sep$, Optional NoTrim As Boolean)
RmvBef = Brk2(S, Sep, NoTrim).S2
End Function

Function BefAt(S, At&)
If At = 0 Then Exit Function
BefAt = Left(S, At - 1)
End Function

Function BefDD$(S)
BefDD = RTrim(BefOrAll(S, "--"))
End Function

Function BefDDD$(S)
BefDDD = RTrim(BefOrAll(S, "---"))
End Function

Function BefMust$(S, Sep$, Optional NoTrim As Boolean)
BefMust = Brk(S, Sep, NoTrim).S1
End Function

Function BefOrAll$(S, Sep$, Optional NoTrim As Boolean)
BefOrAll = Brk1(S, Sep, NoTrim).S1
End Function

Function BefOrAllRev$(S, Sep$)
BefOrAllRev = StrDft(BefRev(S, Sep), Sep$)
End Function

Private Sub BefFstLas__Tst()
Dim S, Fst$, Las$
S = " A_1$ = ""Function ZChunk$(ConstLy$(), IChunk%)"" & _"
Fst = vbDblQ
Las = vbDblQ
Ept = "Function ZChunk$(ConstLy$(), IChunk%)"
GoSub Tst
Exit Sub
Tst:
    Act = BetFstLas(S, Fst, Las)
    C
    Return
End Sub

Function BetFstLas$(S, Fst$, Las$)
BetFstLas = BefRev(Aft(S, Fst), Las)
End Function
Function BetLng(L&, A&, B&) As Boolean
BetLng = A <= L And L <= B
End Function

Function BetStr$(S, S1$, S2$, Optional NoTrim As Boolean, Optional InlMarker As Boolean)
With Brk1(S, S1, NoTrim)
   If .S2 = "" Then Exit Function
   Dim O$: O = Brk1(.S2, S2, NoTrim).S1
   If InlMarker Then O = S1 & O & S2
   BetStr = O
End With
End Function

Private Sub AftBkt__Tst()
Dim A$
A = "(lsk(aa)df lsdkfj) A"
Ept = " A"
GoSub Tst
Exit Sub
Tst:
    Act = AftBkt(A)
    C
    Return
End Sub

Private Sub Bet__Tst()
Dim Ln
Ln = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??       | DATABASE= | ; | ??":            GoSub Tst
Ln = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX | DATABASE= | ; | ??":            GoSub Tst
Ln = "lkjsdf;dkfjl;Data Source=Johnson;lsdfjldf  | Data Source= | ; | Johnson":    GoSub Tst
Exit Sub
Tst:
    Dim FmStr$, ToStr$
    AsgAy AmTrim(SplitVBar(Ln)), Ln, FmStr, ToStr, Ept
    Act = IsBet(Ln, FmStr, ToStr)
    C
    Return
End Sub

Private Sub BetBkt__Tst()
Dim A$
Ept = "1234()567": A = "sdklfjdsf(1234()567)aaa(": GoSub Tst
Ept = "AA":        A = "XXX(AA)XX":                GoSub Tst
Ept = "A$()A":     A = "(A$()A)XX":                GoSub Tst
Ept = "O$()":      A = "(O$()) As X":              GoSub Tst
Ept = "1234()567": A = "sdklfjdsf(1234()567)aaa(": GoSub Tst
Exit Sub

Tst:
    Act = BetBkt(A)
    C
    Return
End Sub

Function BefRevOrAll$(S, Sep$)
Dim P%: P = InStrRev(S, Sep)
If P = 0 Then BefRevOrAll = S: Exit Function
BefRevOrAll = Left(S, P - Len(Sep))
End Function
