Attribute VB_Name = "MxIdeSrcDclCnstNsFt"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcDclCnstNsFt."

Function MdNsFt$(P As VBProject)
MdNsFt = AssPthzP(P) & "MdNs.txt"
End Function
Function MdNsFtP$()
MdNsFtP = MdNsFt(CPj)
End Function
