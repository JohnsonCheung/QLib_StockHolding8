Attribute VB_Name = "MxVbAyNwSno"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Sno"
Const CMod$ = CLib & "MxVbAyNwSno."

Function Sno(N, Optional Fm = 0) As Integer()
Sno = SnoInto(EmpIntAy, N, Fm)
End Function
Function IntSno(N, Optional Fm = 0) As Integer()
IntSno = SnoInto(EmpIntAy, N, Fm)
End Function
Function LngSno(N, Optional Fm = 0) As Long()
LngSno = SnoInto(EmpLngAy, N, Fm)
End Function
Function BytSno(N, Optional Fm = 0) As Byte()
BytSno = SnoInto(EmpBytAy, N, Fm)
End Function
Private Function SnoInto(Into, N, Optional Fm = 0)
Dim O: O = Into: Erase O
Dim J&: For J = Fm To N + Fm - 1
    PushI O, J
Next
SnoInto = O
End Function

Function IxCol(N&, BegIx%) As String() ' Ret
If BegIx < 0 Then Exit Function
IxCol = AmAliR(AddAy(UL(String(NDig(N), "#")), LngSno(N, BegIx)))
End Function
