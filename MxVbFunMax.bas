Attribute VB_Name = "MxVbFunMax"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVbFunMax."

Function MaxLinCntMdn$()
MaxLinCntMdn = MaxLinCntMdnzP(CPj)
End Function

Function MaxLinCntMdnzP$(P As VBProject)
MaxLinCntMdnzP = Mdn(MaxLinCntMd(P))
End Function
Function MaxLinCntMd(P As VBProject) As CodeModule
Dim C As VBComponent, M&, N&, I
For Each C In P.VBComponents
    N = C.CodeModule.CountOfLines
    If N > M Then
        M = N
        Set MaxLinCntMd = C.CodeModule
    End If
Next
End Function
