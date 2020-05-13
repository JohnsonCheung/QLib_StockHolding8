Attribute VB_Name = "MxVbStrJn"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay"
Const CMod$ = CLib & "MxVbStrJn."
Option Base 0

Function Jn$(Ay, Optional Sep$ = "")
Jn = Join(SyzAy(Ay), Sep)
End Function

Function QuoBktJnComma$(Ay)
QuoBktJnComma = QuoBkt(JnComma(Ay))
End Function

Function JnComma$(Ay)
JnComma = Jn(Ay, ",")
End Function

Function JnBq$(Ay)
':Bq: :Chr ! #Back-quo#
JnBq = Jn(Ay, "`")
End Function

Function JnCommaCrLf$(Ay)
JnCommaCrLf = Jn(Ay, "," & vbCrLf)
End Function

Function JnAp$(Sep$, ParamArray Ap())
Dim Av(): If UBound(Ap) > 0 Then Av = Ap
JnAp = Jn(Av, Sep)
End Function

Function JnCommaSpc$(Ay)
JnCommaSpc = Jn(Ay, ", ")
End Function

Function JnCrLf$(Ay)
JnCrLf = Jn(Ay, vbCrLf)
End Function

Function JnCrLfAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnCrLfAp = Jn(Av, vbCrLf)
End Function

Function JnDblCrLf$(Ay)
JnDblCrLf = Jn(Ay, vb2CrLf)
End Function

Function JnDotAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) > 0 Then Av = Ap: JnDotAp = JnDot(Av)
End Function

Function JnDotApNB$(ParamArray Ap())
Dim Av(): If UBound(Ap) > 0 Then Av = Ap: JnDotApNB = JnDotNB(Av)
End Function

Function QuoJnzAsTLn$(Ay)
QuoJnzAsTLn = QuoJn(Ay, " | ", "| * |")
End Function

Function QuoJn$(Ay, Sep$, QuoStr$)
QuoJn = Quo(Jn(Ay, Sep), QuoStr)
End Function

Function QuoJnDot$(Ay)
'Ret : a str joining @Ay and qte with . in front and at end
QuoJnDot = QuoDot(JnDot(Ay))
End Function

Function JnDot$(Ay)
JnDot = Jn(Ay, ".")
End Function

Function JnDotNB$(Ay)
JnDotNB = JnNB(Ay, ".")
End Function

Function JnDollar$(Ay)
JnDollar = Jn(Ay, "$")
End Function

Function JnDblDollar$(Ay)
JnDblDollar = Jn(Ay, "$$")
End Function

Function JnPthSepAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnPthSepAp = JnPthSep(Av)
End Function

Function JnPthSep$(Ay)
JnPthSep = Jn(Ay, PthSep)
End Function

Function JnSemi$(Ay)
JnSemi = Jn(Ay, ";")
End Function

Function JnOr$(Ay)
JnOr = Jn(Ay, " or ")
End Function

Function JnSpc$(Ay)
JnSpc = Jn(Ay, " ")
End Function

Function JnTab$(Ay)
JnTab = Join(Ay, vbTab)
End Function

Function JnVBar$(Ay)
JnVBar = Jn(Ay, "|")
End Function

Function JnVbarSpc$(Ay)
JnVbarSpc = Jn(Ay, " | ")
End Function
