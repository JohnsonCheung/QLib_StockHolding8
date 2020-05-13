Attribute VB_Name = "MxVbStrDblQ"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Str.Quo"
Const CMod$ = CLib & "MxVbStrDblQ."

Private Sub BetDblQ__Tst()
MsgBox QuoSq(BetDblQ("A""XX 1""dd"))
End Sub

Function BetDblQ$(S)
Dim P%, P2%, L%
P = InStr(S, vbDblQ): If P = 0 Then Exit Function
P = P + 1
P2 = InStr(P, S, vbDblQ): If P2 = 0 Then Exit Function
L = P2 - P
BetDblQ = Mid(S, P, L)
End Function
