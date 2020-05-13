Attribute VB_Name = "MxX"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Vb"
Const CMod$ = CLib & "MxX."
Private A$()
Sub ClrXX(): Erase A: End Sub
Sub XBox(S$): X Box(S): End Sub
Sub XEnd(): PushI A, "End": End Sub
Sub XDrs(Drs As Drs): PushIAy A, FmtDrs(Drs): End Sub
Sub XLin(Optional L$): PushI A, L: End Sub

Function XX() As String()
XX = A
Erase A
End Function


Sub XTab(V)
If IsArray(V) Then
    X AmTab(V)
Else
    X vbTab & V
End If
End Sub

Sub X(V)
If IsArray(V) Then
    PushIAy A, V
Else
    PushI A, V
End If
End Sub
