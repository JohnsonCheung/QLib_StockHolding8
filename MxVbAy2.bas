Attribute VB_Name = "MxVbAy2"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Ay.Ay2"
Const CMod$ = CLib & "MxVbAy2."
Type Ay2
    A As Variant
    B As Variant
End Type

Function Ay2(A, B) As Ay2
ChkIsAy A, CSub
ChkIsAy B, CSub
With Ay2
    .A = A
    .B = B
End With
End Function

Function Ay2zAyPfx(Ay, Pfx$) As Ay2
Dim O As Ay2
O.A = NwAy(Ay)
O.B = O.A
Dim S$, I
For Each I In Itr(Ay)
    S = I
    If HasPfx(S, Pfx) Then
        PushI O.B, S
    Else
        PushI O.A, S
    End If
Next
Ay2zAyPfx = O
End Function

Function Ay2zAyN(Ay, N&) As Ay2
Ay2zAyN = Ay2(FstNEle(Ay, N), AeFstNEle(Ay, N))
End Function

Function Ay2Jn(A, B, Sep$) As String()
Dim J&: For J = 0 To Min(UB(A), UB(B))
    PushI Ay2Jn, A(J) & Sep & B(J)
Next
End Function

Function Ay2JnDot(A, B) As String()
Ay2JnDot = Ay2Jn(A, B, ".")
End Function

Function Ay2JnSngQ(A, B) As String()
Ay2JnSngQ = Ay2Jn(A, B, "'")
End Function

Function Ay3FmAyBei(Ay, B As Bei) As Ay3
Ay3FmAyBei = Ay3FmAyBE(Ay, B.Bix, B.Eix)
End Function

Function DyFmAy2(A, B) As Variant()
Dim J&
For J = 0 To Min(UB(A), UB(B))
    PushI DyFmAy2, Array(A(J), B(J))
Next
End Function

Function DrsFmAy2(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2") As Drs
DrsFmAy2 = Drs(Sy(N1, N2), DyFmAy2(A, B))
End Function

Function FmtAyabSpc(AyA, AyB) As String()
FmtAyabSpc = FmtAyab(AyA, AyB, " ")
End Function

Function FmtAyab(A, B, Optional Sep$, Optional FF$ = "Ay1 Ay2") As String()
FmtAyab = FmtS12y(S12yzAyab(A, B), FF)
End Function

Function FmtAyabNEmpB(A, B, Optional Sep$ = " ") As String()
Dim J&, O$()
For J = 0 To UB(A)
    If Not IsEmp(B(J)) Then
        Push O, A(J) & Sep & B(J)
    End If
Next
FmtAyabNEmpB = O
End Function

Sub AsgAyaReSzMax(A, B, OA, OB)
OA = A
OB = B
ResiMax OA, OB
End Sub

Function DyzAy2(A, B) As Variant()
ChkSamSi A, B, CSub
Dim I, J&: For Each I In Itr(A)
    PushI DyzAy2, Array(I, B(J))
    J = J + 1
Next
End Function
