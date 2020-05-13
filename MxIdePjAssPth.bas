Attribute VB_Name = "MxIdePjAssPth"
Option Explicit
Option Compare Text
Const CNs$ = "Pj.AssPth"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxIdePjAssPth."

Sub BrwAssPthP()
BrwPth AssPthP
End Sub

Function AssPthzM$(M As CodeModule)
AssPthzM = AssPthzP(PjzM(M))
End Function

Function AssPthzP$(P As VBProject)
AssPthzP = EnsPth(AssPth(Pjf(P)))
End Function

Function AssPthP$()
AssPthP = AssPthzP(CPj)
End Function

Sub EnsAssPthP()
EnsAssPthzP CPj
End Sub

Private Sub EnsAssPthzP(P As VBProject)
EnsPth AssPthzP(P)
End Sub
