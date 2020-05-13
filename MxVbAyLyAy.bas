Attribute VB_Name = "MxVbAyLyAy"
Option Compare Text
Option Explicit
Const CNs$ = "Vb.Dta"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbAyLyAy."

Sub VcLyAy(LyAy())
Vc FmtLyAy(LyAy)
End Sub

Sub BrwLyAy(LyAy())
BrwAy FmtLyAy(LyAy)
End Sub

Function FmtLyAy(LyAy()) As String()
Dim O$(), H$
Dim W%: W = WdtzLyAy(LyAy)
H = Quo(Dup("-", W + 2), "|")
Dim J%: For J = 0 To UB(LyAy)
    PushI FmtLyAy, H
    PushIAy FmtLyAy, AmQuoAli(LyAy(J), W, "| * |")
Next
PushI FmtLyAy, H
End Function
Function AmQuoAli(Ay, W%, QuoStr$) As String()
Dim Q As S12: Q = BrkQuo(QuoStr)
Dim I: For Each I In Itr(Ay)
    PushI AmQuoAli, Q.S1 & Ali(I, W) & Q.S2
Next
End Function
Function WdtzLyAy%(LyAy())
Dim O%
Dim Sy: For Each Sy In Itr(LyAy)
    O = Max(O, WdtzAy(Sy))
Next
WdtzLyAy = O
End Function
