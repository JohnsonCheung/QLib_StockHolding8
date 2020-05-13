Attribute VB_Name = "JMxFmt"
Option Compare Text
Const CMod$ = CLib & "JMxFmt."
#If False Then
Option Explicit
Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
FmtQQ = FmtQQAv(QQVbl, Av)
End Function

Function FmtQQAv$(QQVbl$, Av())
Dim O$: O = Replace(QQVbl, "|", vbCrLf)
Dim P&: P = 1
Dim I: For Each I In Av
    P = InStr(P, O, "?")
    If P = 0 Then Exit For
    O = Left(O, P - 1) & Replace(O, "?", I, Start:=P, Count:=1)
    P = P + Len(I)
Next
FmtQQAv = O
End Function
#End If
