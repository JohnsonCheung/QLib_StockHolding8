Attribute VB_Name = "MxIdeSrcRmvSubZ"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcRmvSubZ."
Sub RmvSubZM()
RmvSubYYM CMd
End Sub
Sub RmvSubYYM(M As CodeModule)
DltMth M, "Z"
End Sub
Sub RmvSubZP()
RmvSubYYP CPj
End Sub
Sub RmvSubYYP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    RmvSubYYM C.CodeModule
Next
End Sub
