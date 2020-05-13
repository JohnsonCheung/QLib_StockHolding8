Attribute VB_Name = "MxIdeSrcDclUdtIx"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcDclUdtIx."

Function UdtBix&(Dcl$(), FmIx%) ' Ret -1 if not found
Dim J%: For J = FmIx To UB(Dcl)
   If IsUdtln(Dcl(J)) Then UdtBix = J: Exit Function
Next
UdtBix = -1
End Function

Private Function UdtBixzN&(Dcl$(), Udtn$) ' Ret -1 if not found
Dim J%: For J = 0 To UB(Dcl)
   If UdtnzL(Dcl(J)) = Udtn Then UdtBixzN = J: Exit Function
Next
UdtBixzN = -1
End Function

Private Function UdtEix%(Dcl$(), Bix%) 'Return -1 if Bix is < 0 or the Eix of [End Type] line
If Bix < 0 Then UdtEix = -1: Exit Function
Dim J%: For J = Bix To UB(Dcl)
    If HasSubStr(Dcl(J), "End Type") Then
        UdtEix = J
        Exit Function
    End If
Next
Thw CSub, "No End Type is found", "Bix Dcl", Bix, Dcl
End Function
Function UdtBei(Dcl$(), Udtn$) As Bei
Dim B%: B = UdtBixzN(Dcl, Udtn)
UdtBei = Bei(B, UdtEix(Dcl, B))
End Function

Function UdtBeiy(Dcl$()) As Bei()
Dim FmIx%, B%, E%, J%
Again:
    LoopTooMuch CSub, J
    B = UdtBix(Dcl, FmIx): If B < 0 Then Exit Function
    E = UdtEix(Dcl, B):  PushBei UdtBeiy, Bei(B, E)
    FmIx = E + 1
    GoTo Again
End Function

Private Sub IsUdtln__Tst()
Dim O$()
Dim L: For Each L In DclP
    If IsUdtln(L) Then PushI O, L
Next
BrwAy O
End Sub

Private Function IsUdtln(Ln) As Boolean
IsUdtln = ShfTermX(RmvMdy(Ln), "Type")
End Function
