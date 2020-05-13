Attribute VB_Name = "MxIdePjLen"
Option Explicit
Option Compare Text
Const CNs$ = "Pj.Prp"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdePjLen."

Function MdLenM&()
MdLenM = MdLen(CMd)
End Function
Function MdLen&(M As CodeModule)
MdLen = Len(Srcl(M))
End Function

Function PjLen&(P As VBProject)
Dim O&, C As VBComponent
For Each C In P.VBComponents
    O = O + MdLen(C.CodeModule)
Next
PjLen = O
End Function

Function PjLenP&()
PjLenP = PjLen(CPj)
End Function
