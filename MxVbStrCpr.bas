Attribute VB_Name = "MxVbStrCpr"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str.Cpr"
Const CMod$ = CLib & "MxVbStrCpr."

Sub CprLines(A$, B$, Optional N12$ = "A B", Optional Hdr$)
Brw FmtCprLines(A, B, N12, Hdr)
End Sub

'---=====================================

Private Sub FmtCprLines__Tst()
Dim A$, B$
A = LineszVbl("AAAAAAA|bbbbbbbb|cc|dd")
B = LineszVbl("AAAAAAA|bbbbbbbb |cc")
GoSub Tst
Exit Sub
Tst:
    Act = FmtCprLines(A, B)
    Brw Act
    Return

End Sub

Function FmtCprLines(A$, B$, Optional N12$ = "A B", Optional Hdr$) As String()
If A = B Then PushI FmtCprLines, "Two lines are equal.  N12=[" & N12 & "]"
Dim AA$(): AA = SplitCrLf(A)
Dim BB$(): BB = SplitCrLf(B)
Dim N1$, N2$: AsgTRst N12, N1, N2
Dim NIxDig%: NIxDig = Len(Max(Si(AA), Si(BB)))
Dim H$(): H = W1Hdr(AA, BB, N1, N2, Hdr)
Dim L$(): L = W1Cpr(AA, BB, N1, N2, NIxDig)
Dim R$(): R = W1Rst(AA, BB, N1, N2)
FmtCprLines = AddSyAp(H, L, R)
End Function

Private Function W1Hdr(A$(), B$(), N1$, N2$, Hdr$) As String()  ' The [Hdr] part
Dim O$()
PushI O, FmtQQ("LinesCnt=? (?)", Si(A), N1)
PushI O, FmtQQ("LinesCnt=? (?)", Si(B), N2)
W1Hdr = O
End Function

Private Function W1Cpr(A$(), B$(), N1$, N2$, NIxDig%) As String() ' The [Cpr] part
Dim J&: For J = 0 To Min(UB(A), UB(B))
    PushI W1Cpr, W1Lin(A(J), B(J), J, NIxDig)
Next
End Function

Private Function W1Lin(A$, B$, Ix&, NIxDig%) As String() ' The [Lin] which will be 1 line if same or 2 lines if dif
PushI W1Lin, W1FmtLn(Ix, NIxDig, A)
If A = B Then Exit Function
PushI W1Lin, W1FmtLnSpc(NIxDig, B)
End Function

Private Function W1Rst(A$(), B$(), N1$, N2$) As String() ' The [Rst] part
Dim NA&, NB&, N&: NA = Si(A): NB = Si(B): N = Max(NA, NB)
If NA = NB Then Exit Function
Dim MoreNm$, LessNm$, NLn&
    If NA > NB Then
        MoreNm = N1
        LessNm = N2
    Else
        MoreNm = N2
        LessNm = N1
    End If
    NLn = Abs(NA - NB)

Dim O$()
    PushI O, FmtQQ("-- ? has more ? lines then ? -------", MoreNm, NLn, LessNm) '<===
    Dim Large$()
        If NA > NB Then
            Large = A
        Else
            Large = B
        End If
    Dim J&
    Dim NIxDig%: NIxDig = Len(N)
    For J = Min(NA, NB) To N - 1
        PushI W1Rst, W1FmtLn(J, NIxDig, Large(J))
    Next
W1Rst = O
End Function

Private Function W1FmtLn$(Ix&, NIxDig%, Ln$)
W1FmtLn = FmtQQ("? ?", AliR(Ix + 1, NIxDig), Ln)
End Function

Private Function W1FmtLnSpc$(NIxDig%, Ln)
W1FmtLnSpc = Space(NIxDig + 1) & Ln
End Function

'---=====================================
Sub ChkIsEqStr(A$, B$, Optional N12$ = "A B", Optional Hdr$)
If A = B Then Exit Sub
Brw FmtCprStr(A, B, N12, Hdr)
End Sub

Function FmtCprStr(A$, B$, Optional N12$ = "A B", Optional Hdr$) As String()
Dim N1$, N2$
AsgTRst N12, N1, N2
If A = B Then
    PushI FmtCprStr, FmtQQ("Str(?) = Str(?).  Len(?)", N1, N2, Len(A))
    Exit Function
End If
Select Case True
Case IsLines(A), IsLines(B): FmtCprStr = FmtCprLines(A, B, N12)
Case Else: FmtCprStr = W2FmtCprStr(A, B, N1, N2, Hdr)
End Select
End Function

Private Function W2FmtCprStr(A$, B$, N1$, N2$, Hdr$) As String() '
Dim P&: P = DifPos(A, B)
    Dim L1&, L2&: L1 = Len(A): L2 = Len(B)
Dim O$()
    PushIAy O, Box(Hdr)
    PushI O, FmtQQ("Str1 (Len / Nm): ? / ?", AliR(L1, 6), N1)
    PushI O, FmtQQ("Str2 (Len / Nm): ? / ?", AliR(L2, 6), N2)
    PushI O, FmtQQ("Dif at position: ?", P)
    PushIAy O, Lbl123(Max(L1, L2))
    PushI O, A
    PushI O, B
    PushI O, Space(P - 1) & "^"
W2FmtCprStr = O
End Function
