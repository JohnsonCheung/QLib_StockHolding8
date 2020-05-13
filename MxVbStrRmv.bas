Attribute VB_Name = "MxVbStrRmv"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbStrRmv."
Const CNs$ = "Str"


Function RmvDotComma$(S)
RmvDotComma = Replace(Replace(S, ",", ""), ".", "")
End Function
Function Rmv2Dash$(S)
Rmv2Dash = RTrim(RmvAft(S, "--"))
End Function

Function Rmv3Dash$(S)
Rmv3Dash = RTrim(RmvAft(S, "---"))
End Function

Function Rmv3T$(S)
Rmv3T = RmvTT(RmvT1(S))
End Function

Function RmvP12$(S, P As C12) ' Rmv the str point by @Bet inclusive
If IsEmpC12(P) Then RmvP12 = S: Exit Function
RmvP12 = Left(S, P.C1 - 1) & Mid(S, P.C2 + 1)
End Function

Function RmvAft$(S, Sep$)
RmvAft = Brk1(S, Sep, NoTrim:=True).S1
End Function

Function RmvDblSpc$(S) ' Rpl more than one spc to one.
Dim O$: O = S
While HasSubStr(O, "  ")
    O = Replace(O, "  ", " ")
Wend
RmvDblSpc = O
End Function

Function RmvFstChr$(S)
RmvFstChr = Mid(S, 2)
End Function

Function RmvFst2Chr$(S)
RmvFst2Chr = Mid(S, 3)
End Function

Function RmvFstLasChr$(S)
RmvFstLasChr = RmvFstChr(RmvLasChr(S))
End Function

Function RmvFstNChr$(S, Optional N% = 1)
RmvFstNChr = Mid(S, N + 1)
End Function

Function RmvFstNonLetter$(S)
If IsAscLetter(Asc(S)) Then
    RmvFstNonLetter = S
Else
    RmvFstNonLetter = RmvFstChr(S)
End If
End Function

Function RmvLas2Chr$(S)
RmvLas2Chr = RmvLasNChr(S, 2)
End Function

Function RmvLasChr$(S)
RmvLasChr = RmvLasNChr(S, 1)
End Function

Function RmvLasTwoChr$(S)
RmvLasTwoChr = RmvLasNChr(S, 2)
End Function

Function RmvLasNChr$(S, N%)
Dim L&: L = Len(S) - N: If L <= 0 Then Exit Function
RmvLasNChr = Left(S, L)
End Function

Function RmvNm$(S)
Dim O%
If Not IsAscFstNmChr(Asc(FstChr(S))) Then GoTo X
For O = 1 To Len(S)
    If Not IsAscNmChr(Asc(Mid(S, O, 1))) Then GoTo X
Next
X:
    If O > 0 Then RmvNm = Mid(S, O): Exit Function
    RmvNm = S
End Function

Function RmvSqBktzSy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI RmvSqBktzSy, RmvSqBkt(I)
Next
End Function
Function RmvSqBkt$(S)
If Not HasSqBkt(S) Then RmvSqBkt = S: Exit Function
RmvSqBkt = RmvFstLasChr(S)
End Function

Function RmvPfxAll$(S, Pfx$)
Dim O$: O = S
Dim J%
While HasPfx(O, Pfx)
    LoopTooMuch CSub, J
    O = RmvPfx(O, Pfx)
Wend
RmvPfxAll = O
End Function

Function RmvPfx$(S, Pfx$, Optional C As VbCompareMethod = vbTextCompare)
If HasPfx(S, Pfx) Then RmvPfx = Mid(S, Len(Pfx) + 1) Else RmvPfx = S
End Function

Function RmvPfxSy$(S, PfxSy$(), Optional C As VbCompareMethod = vbTextCompare)
Dim Pfx$, I
For Each I In PfxSy
    Pfx = I
    If HasPfx(S, Pfx, C) Then RmvPfxSy = RmvPfx(S, Pfx, C): Exit Function
Next
RmvPfxSy = S
End Function
Function RmvPfxSpc$(S, Pfx$)
If Not HitPfxSpc(S, Pfx) Then RmvPfxSpc = S: Exit Function
RmvPfxSpc = LTrim(Mid(S, Len(Pfx) + 2))
End Function
Function RmvPfxSySpc$(S, PfxSy$())
Dim I, Pfx$
For Each I In PfxSy
    Pfx = I
    If HitPfxSpc(S, Pfx) Then
        RmvPfxSySpc = LTrim(Mid(S, Len(Pfx) + 2))
        Exit Function
    End If
Next
RmvPfxSySpc = S
End Function

Function RmvBkt$(S)
RmvBkt = RmvSfxzBkt(S)
End Function

Function RmvSfxzBkt$(S)
RmvSfxzBkt = RmvSfx(S, "()")
End Function

Function RmvSfxDot$(S)
RmvSfxDot = RmvSfx(S, ".")
End Function

Function RmvSfx$(S, Sfx$, Optional B As VbCompareMethod = vbBinaryCompare)
If HasSfx(S, Sfx, B) Then RmvSfx = Left(S, Len(S) - Len(Sfx)) Else RmvSfx = S
End Function

Function RmvSngQuo$(S)
If Not IsSngQuoted(S) Then RmvSngQuo = S: Exit Function
RmvSngQuo = RmvFstLasChr(S)
End Function

Function RmvTerm$(S, Term$)
Dim T$: T = T1(S)
If T = Term Then
    RmvTerm = Mid(S, Len(T) + 1)
End If
    RmvTerm = S
End Function

Function RmvT1$(S)
Dim L$: L = S: ShfTerm L
RmvT1 = L
End Function

Function RmvTT$(S)
RmvTT = RmvT1(RmvT1(S))
End Function

Private Sub RmvT1__Tst()
Ass RmvT1("  df dfdf  ") = "dfdf"
End Sub

Private Sub RmvNm__Tst()
Dim Nm$
Nm = "lksdjfsd f"
Ept = " f"
GoSub Tst
Exit Sub
Tst:
    Act = RmvNm(Nm)
    C
    Return
End Sub

Private Sub RmvPfx__Tst()
Ass RmvPfx("aaBB", "aa") = "BB"
End Sub

Private Sub RmvPfxSy__Tst()
Dim S, PfxSy$()
PfxSy = SyzSS("Z_ Z_"): Ept = "ABC"
S = "Z_ABC": GoSub Tst
S = "Z_ABC": GoSub Tst
Exit Sub
Tst:
    Act = RmvPfxSy(S, PfxSy)
    C
    Return
End Sub

Function RmvCr$(S)
RmvCr = Replace(S, vbCr, "")
End Function
Function RmvEndDig$(S)
Dim J&: For J = Len(S) To 1 Step -1
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then
        RmvEndDig = Left(S, J)
        Exit Function
    End If
Next
RmvEndDig = Left(S, J)
End Function
