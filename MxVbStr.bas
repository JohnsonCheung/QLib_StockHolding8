Attribute VB_Name = "MxVbStr"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVbStr."
Function SzTrue$(B As Boolean, S)
If B Then SzTrue = S
End Function
Function SzFalse$(B As Boolean, S)
If B = False Then SzFalse = S
End Function

Function Pad0$(N, NDig)
Pad0 = Format(N, Dup("0", NDig))
End Function

Sub BrwStr(S, Optional Fnn$, Optional UseVc As Boolean)
Dim T$: T = TmpFt("BrwStr", Fnn$)
WrtStr S, T
BrwFt T, UseVc
End Sub

Sub VcStr(S, Optional Fnn$)
BrwStr S, Fnn, UseVc:=True
End Sub

Function StrDft$(S, Dft)
StrDft = IIf(S = "", Dft, S)
End Function

Function Dup$(S, N)
Dim O$, J&
For J = 0 To N - 1
    O = O & S
Next
Dup = O
End Function

Function HasSfxAs(S, AsSfxSy$()) As Boolean
Dim I, Sfx$
For Each I In Itr(AsSfxSy)
    Sfx = I
    If HasSfx(S, Sfx) Then HasSfxAs = True: Exit Function
Next
End Function

Function HasPfxAs(S, AsPfxSy$()) As Boolean
Dim I, Pfx$
For Each I In Itr(AsPfxSy)
    Pfx = I
    If HasPfx(S, Pfx) Then HasPfxAs = True: Exit Function
Next
End Function

Sub EdtStr(S, Ft)
WrtStr S, Ft, OvrWrt:=True
Brw Ft
End Sub

Function IsDigStr(S) As Boolean
Dim J&
For J = 1 To Len(S)
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then Exit Function
Next
IsDigStr = True
End Function

Function LeftOrAll(S, L%)
If L <= 0 Then
    LeftOrAll = S
Else
    LeftOrAll = Left(S, L)
End If
End Function

Function DblQPosy(S) As Integer(): DblQPosy = SubStrPosy(S, vbDblQ): End Function

Function SubStrPosy(S, SubStr$) As Integer()
Dim P%: P = 1
Dim M%, J%, L%
L = Len(SubStr)
Again:
    LoopTooMuch CSub, J
    M = InStr(P, S, SubStr): If M = 0 Then Exit Function
    P = M + L
    PushI SubStrPosy, M
    GoTo Again
End Function

Function IsIndtln(Ln) As Boolean ' ret True if after Trim fst chr is a ""
IsIndtln = LTrim(FstChr(Ln)) = ""
End Function

Function IsHdrln(Ln) As Boolean ' ret Not IsIndtln
IsHdrln = Not IsIndtln(Ln)
End Function

Function Ali$(V, W%)
Dim S: S = V
If IsStr(V) Then
    Ali = AliL(S, W)
Else
    Ali = AliR(S, W)
End If
End Function

Function AliL$(S, W)
Dim L%: L = Len(S)
If L >= W Then
    AliL = S
Else
    AliL = S & Space(W - Len(S))
End If
End Function

Function AliR$(S, W)
Dim L%: L = Len(S)
If W > L Then
    AliR = Space(W - L) & S
Else
    AliR = S
End If
End Function

Function AliRzT1(Ly$()) As String()
Dim T1$(), Rst$()
AsgAmT1RstAy Ly, T1, Rst
T1 = AmAliR(T1)
Dim J&: For J = 0 To UB(T1)
    PushI AliRzT1, T1(J) & " " & Rst(J)
Next
End Function

Function TrimWhite$(A)
TrimWhite = TrimWhiteL(TrimWhiteL(A))
End Function

Function TrimWhiteL$(A)
Dim J%
    For J = 1 To Len(A)
        If Not IsWhiteChr(Mid(A, J, 1)) Then Exit For
    Next
TrimWhiteL = Left(A, J)
End Function

Function TrimWhiteR$(S)
Dim J%
    Dim A$
    For J = Len(S) To 1 Step -1
        If Not IsWhiteChr(Mid(S, J, 1)) Then Exit For
    Next
    If J = 0 Then Exit Function
TrimWhiteR = Mid(S, J)
End Function

Function TabN$(N%)
TabN = Space(4 * N)
End Function

Function Rpl$(S, SubStr$, By$, Optional Ith% = 1)
Dim P&: P = InStrWiIthSubStr(S, SubStr, Ith)
If P = 0 Then Rpl = S: Exit Function
Rpl = Replace(S, SubStr, By, P, 1)
End Function

Function VzS(S, T As VbVarType)
Dim O
Select Case True
Case T = vbBoolean: O = CBool(S)
End Select
End Function
