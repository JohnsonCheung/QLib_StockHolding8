Attribute VB_Name = "MxIdeSrcFfn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcFfn."

Function SrcFfnM$(): SrcFfnM = SrcFfn(CCmp): End Function
Private Sub SrcFfn__Tst(): Vc SrcFfnAyP: End Sub
Function SrcFfnAyP() As String(): SrcFfnAyP = SrcFfnyzP(CPj): End Function
Function SrcFfn$(A As VBComponent): SrcFfn = SrcPthzCmp(A) & SrcFn(A): End Function
Function SrcFn$(A As VBComponent): SrcFn = A.Name & ".bas": End Function
Function SrcFfnzMdn$(Mdn$): SrcFfnzMdn = SrcFfn(Cmp(Mdn)): End Function
Function SrcFfnzM$(M As CodeModule): SrcFfnzM = SrcFfn(M.Parent): End Function
Function SrcFfnyP(): SrcFfnyP = SrcFfnyzP(CPj): End Function
Function SrcFfnyzP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI SrcFfnyzP, SrcFfn(C)
Next
End Function

Function ExtzCmpTy$(A As vbext_ComponentType)
Dim O$
Select Case A
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case vbext_ct_MSForm: O = ".cls"
Case Else: Raise "SrcExt: Unexpected Md_CmpTy.  Should be [Class or Module or Document]"
End Select
ExtzCmpTy = O
End Function

