Attribute VB_Name = "JMxStrRmv"
Option Compare Text
Const CMod$ = CLib & "JMxStrRmv."
#If False Then
Option Explicit

Function RmvFstLasChr$(S)
RmvFstLasChr = RmvFstChr(RmvLasChr(S))
End Function

Function RmvFstChr$(S)
RmvFstChr = Mid(S, 2)
End Function

Function RmvLasChr$(S)
RmvLasChr = RmvLasNChr(S, 1)
End Function

Function RmvLasNChr$(S, N%)
Dim L&: L = Len(S) - N: If L <= 0 Then Exit Function
RmvLasNChr = Left(S, L)
End Function

Function RmvSfx$(S, Sfx$, Optional B As VbCompareMethod = vbBinaryCompare)
If HasSfx(S, Sfx, B) Then RmvSfx = Left(S, Len(S) - Len(Sfx)) Else RmvSfx = S
End Function

#End If
