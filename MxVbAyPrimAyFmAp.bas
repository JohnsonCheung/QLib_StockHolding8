Attribute VB_Name = "MxVbAyPrimAyFmAp"
Option Explicit
Option Compare Text
Const CNs$ = "Ay.New"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbAyPrimAyFmAp."

Function IntAy(ParamArray Ap()) As Integer()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
IntAy = IntozAy(EmpIntAy, Av)
End Function

Function BoolAy(ParamArray Ap()) As Boolean()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
BoolAy = IntozAy(EmpBoolAy, Av)
End Function

Function LngAy(ParamArray Ap()) As Long()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
LngAy = IntozAy(EmpLngAy, Av)
End Function

Function SngAy(ParamArray Ap()) As Single()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
SngAy = IntozAy(SngAy, Av)
End Function

Function DteAy(ParamArray Ap()) As Date()
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
DteAy = IntozAy(DteAy, Av)
End Function
