Attribute VB_Name = "MxIdeMdnByMthn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdnByMthn."

Function MdnyzMthn(Mthn) As String()
MdnyzMthn = MdnyzMthnPj(Mthn, CPj)
End Function

Function MdnyzMthnPj(Mthn, P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If HasMthzM(C.CodeModule, Mthn) Then PushI MdnyzMthnPj, C.Name
Next
End Function
