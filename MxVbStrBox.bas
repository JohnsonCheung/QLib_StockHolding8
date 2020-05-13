Attribute VB_Name = "MxVbStrBox"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrBox."

Function BoxzLines(Lines, Optional C$ = "*") As String()
BoxzLines = BoxzLy(SplitCrLf(Lines))
End Function

Function BoxzLy(Ly$(), Optional C$ = "*") As String()
If Si(Ly) = 0 Then Exit Function
Dim W%, L$, I
W = WdtzAy(Ly)
L = Quo(Dup("-", W), "|-*-|")
PushI BoxzLy, L
For Each I In Ly
    PushI BoxzLy, "| " & AliL(I, W) & " |"
Next
PushI BoxzLy, L
End Function

Function BoxzS(S$, Optional C$ = "*") As String()
If Trim(S) = "" Then Exit Function
Dim H$: H = Dup(C, Len(S) + 6)
PushI BoxzS, H
PushI BoxzS, C & C & " " & S & " " & C & C
PushI BoxzS, H
End Function

Function AddTitAy(Tit$, Ay) As String() ' Adding a boxing of title and following an Ay
If Tit = "" Then AddTitAy = Ay: Exit Function
AddTitAy = AddAy(Box(Tit), Ay)
End Function

Function Box(V, Optional C$ = "*") As String()
If V = "" Then Exit Function
If IsStr(V) Then
    If V = "" Then
        Exit Function
    End If
End If
Select Case True
Case IsLines(V): Box = BoxzLines(V, C)
Case IsStr(V):   Box = BoxzS(CStr(V), C)
Case IsSy(V):    Box = BoxzLy(CvSy(Sy), C)
Case IsArray(V): Box = BoxzAy(V)
Case Else:       Box = BoxzS(CStr(V), C)
End Select
End Function

Function Boxl$(V, Optional C$ = "*") 'vbCrLf is always at end
If V = "" Then Exit Function
Boxl = JnCrLf(Box(V, C)) & vbCrLf
End Function

Function BoxzFny(Fny$()) As String()
If Si(Fny) = 0 Then Exit Function
Const S$ = " | ", Q$ = "| * |"
Const LS$ = "-|-", LQ$ = "|-*-|"
Dim L$, H$, Ay$(), J%
    ReDim Ay(UB(Fny))
    For J = 0 To UB(Fny)
        Ay(J) = Dup("-", Len(Fny(J)))
    Next
L = Quo(Jn(Fny, S), Q)
H = Quo(Jn(Ay, LS), LQ)
BoxzFny = Sy(H, L, H)
End Function

Function BoxzAy(Ay) As String()
If Si(Ay) = 0 Then Exit Function
Dim W%: W = WdtzAy(Ay)
Dim H$: H = "|" & Dup("-", W + 2) & "|"
PushI BoxzAy, H
Dim I: For Each I In Ay
    PushI BoxzAy, "| " & AliL(I, W) + " |"
Next
PushI BoxzAy, H
End Function
