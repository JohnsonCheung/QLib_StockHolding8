Attribute VB_Name = "MxVbAyMgeLy"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ly"
Const CMod$ = CLib & "MxVbAyMgeLy."
Function FmtStrColAp(ParamArray StrColAp()) As String()
Dim Av(): Av = StrColAp
FmtStrColAp = FmtStrColy(StrColy(Av), " ")
End Function
Function FmtStrColy(A As StrColy, Optional Sep$ = " ") As String()
Dim C(): C = A.Coly
Dim UCol%: UCol = UB(C): If UCol = -1 Then Exit Function
FmtStrColy = C(0)
Dim J%: For J = 1 To UCol
    FmtStrColy = FmtStrCol(FmtStrColy, CvSy(C(J)), Sep)
Next
End Function
Function FmtStrCol(StrCol1$(), StrCol2$(), Optional Sep$ = " ") As String()
Dim U1&, U2&, U&: U1 = UB(StrCol1): U2 = UB(StrCol2): U = Max(U1, U2): If U = -1 Then Exit Function
Dim A$(): A = StrCol1: If U2 > U1 Then ReDim Preserve A(U2)
Dim B$(): B = StrCol2: If U1 > U2 Then ReDim Preserve B(U1)
A = AmAli(A)
B = AmAli(B)
Dim O$(): ReDim O(U)
Dim J&: For J = 0 To U
    O(J) = A(J) & Sep & B(J)
Next
FmtStrCol = O
End Function
