Attribute VB_Name = "MxIdeSrcDclDrs"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Drs"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcDclDrs."
Public Const DclFF$ = "Pjn Mdn Dcll"

Function DclDrsP() As Drs
DclDrsP = DclDrszP(CPj)
End Function

Function DclDrszP(P As VBProject) As Drs
Dim Dy(), Pjn$
Pjn = P.Name
Dim C As VBComponent: For Each C In P.VBComponents
    PushI Dy, Array(Pjn, C.Name, Dcll(C.CodeModule))
Next
DclDrszP = DrszFF(DclFF, Dy)
End Function
