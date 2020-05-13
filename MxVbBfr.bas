Attribute VB_Name = "MxVbBfr"
Option Compare Text
Option Explicit
Private X$()
Const CLib$ = "QVb."
Const CNs$ = "Fmt"
Const CMod$ = CLib & "MxVbBfr."
Private A$()

Sub BfrClr()
Erase A
End Sub
Sub BfrLin()
PushI A, ""
End Sub

Sub BfrV(Optional V)
If IsEmpty(V) Then PushI A, "": Exit Sub
PushIAy A, Fmt(V)
End Sub

Sub BfrBox(S$, Optional C$ = "*")
PushIAy A, Box(S, C)
End Sub

Sub BfrULin(S$, Optional ULinChr$ = "-")
PushI A, S
PushI A, Dup(FstChr(ULinChr), Len(S))
End Sub

Sub BfrBrw()
BrwAy A
End Sub

Function BfrLy() As String()
BfrLy = A
End Function

Function BfrLines$()
BfrLines = JnCrLf(BfrLy)
End Function

Sub BfrTab(V)
If IsArray(V) Then
    BfrV AmTab(V)
Else
    BfrV vbTab & V
End If
End Sub
