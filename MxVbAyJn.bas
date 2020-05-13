Attribute VB_Name = "MxVbAyJn"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay.Op"
Const CMod$ = CLib & "MxVbAyJn."

Function JnSpcApNB$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnSpcApNB = JnSpc(SyzAyNB(Av))
End Function

Function JnDollarAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnDollarAp = JnDollar(Av)
End Function

Function JnVbarAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnVbarAp = JnVBar(Av)
End Function

Function JnVbarApSpc$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnVbarApSpc = JnVbarSpc(Av)
End Function

Function JnSpcAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnSpcAp = JnSpc(Av)
End Function

Function JnTabAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnTabAp = JnTab(Av)
End Function

Function JnSemiColonAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnSemiColonAp = JnSemi(AeEmpEle(Av))
End Function

Function JnCommaSpcFF$(FF$)
JnCommaSpcFF = JnQSqCommaSpc(Termy(FF))
End Function
