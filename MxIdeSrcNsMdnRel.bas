Attribute VB_Name = "MxIdeSrcNsMdnRel"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxIdeSrcNsMdnRel."

Private Sub NsMdnnRel__Tst()

End Sub


Function NsMdnnAyzP(P As VBProject) As String()

Dim C As VBComponent: For Each C In P.VBComponents
'    PushS12 O, NsMdnnLin(CNsv(Dcl(C.CodeModule)), C.Name)
Next
'S12yoNsqMdnzP = O
End Function

Function ReloNsqMdnzP(P As VBProject) As Dictionary
'Set ReloNsqMdnzP = RelzS12y(S12yoNsqMdnzP(P))
End Function
