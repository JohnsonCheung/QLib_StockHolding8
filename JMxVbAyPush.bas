Attribute VB_Name = "JMxVbAyPush"
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "JMxVbAyPush."
#If False Then
Option Explicit


Sub Push(O, M)
Dim N&: N = Si(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Sub PushNB(O$(), S$)
If Trim(S) <> "" Then PushS O, S
End Sub

Sub PushAy(OAy, Ay)
Dim I: For Each I In Itr(Ay)
    PushI OAy, I
Next
End Sub
Sub PushI(Ay, I)
Dim N&: N = Si(Ay)
ReDim Preserve Ay(N)
Ay(N) = I
End Sub
#End If
