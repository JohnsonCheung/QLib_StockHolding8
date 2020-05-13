Attribute VB_Name = "MxVbStrLy"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbStrLy."

Function FmtStrColAv(StrColAv(), Optional Sep$ = " ")
Dim UCol%: UCol = UB(StrColAv): If UCol = -1 Then Exit Function
Dim URow&: URow = UB(StrColAv(0))
Dim R&: For R = 0 To URow
    Dim Dr$(): Dr = DrzSyAv(StrColAv, R, UCol)
    PushI FmtStrColAv, Jn(Dr, Sep)
Next
End Function

Function AddAliStrCol(StrCol1$(), StrCol2$()) As String()
Dim N1&: N1 = Si(StrCol1): If N1 = 0 Then AddAliStrCol = StrCol2: Exit Function
Dim N2&: N2 = Si(StrCol2): If N2 = 0 Then AddAliStrCol = StrCol1: Exit Function
If N1 <> N2 Then Thw CSub, "Rows of StrCol1 & StrCol2 are dif", "StrCol1-Rows StrCol2-Row", N1, N2
Dim A$(): A = AmAli(StrCol1)
Dim J&: For J = 0 To UB(A)
    PushI AddAliStrCol, A(J) & " " & StrCol2(J)
Next
End Function

Function AddAliStrColAp(ParamArray LyAp()) As String()
Dim LyAv(): LyAv = LyAp
AddAliStrColAp = AddAliStrColAv(LyAv)
End Function

Function AddAliStrColAv(StrColAv()) As String()
If Si(StrColAv) = 0 Then Exit Function
ChkIsSyAv StrColAv, CSub
ChkIsSamSiSyAv StrColAv, CSub
Dim O$(): O = StrColAv(0)
Dim J%: For J = 1 To UB(StrColAv)
    Dim StrCol$(): StrCol = StrColAv(J)
    O = AddAliStrCol(O, StrCol)
Next
AddAliStrColAv = O
End Function
