Attribute VB_Name = "MxIdePjCnt"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CNs$ = "Cnt"
Const CMod$ = CLib & "MxIdePjCnt."

Sub CntCmpP()
BrwDrs CmpCntDrsP
End Sub

Sub CntCmpzP(P As VBProject)
Brw FmtDrsV(CmpCntDrszP(P))
End Sub

Function CmpCntDrsP() As Drs
CmpCntDrsP = CmpCntDrszP(CPj)
End Function

Function CmpCntDrszP(P As VBProject) As Drs
CmpCntDrszP = DrszFF("Pj Tot Mod Cls Doc Frm Oth", Av(CmpCntDrzP(P)))
End Function

Function CmpCntDrzP(P As VBProject) As Variant()
Dim NCls%, NDoc%, NFrm%, NMod%, NOth%, NTot%
Dim C As VBComponent
For Each C In P.VBComponents
    Select Case C.Type
    Case vbext_ct_ClassModule:  NCls = NCls + 1
    Case vbext_ct_Document:     NDoc = NDoc + 1
    Case vbext_ct_MSForm:       NFrm = NFrm + 1
    Case vbext_ct_StdModule:    NMod = NMod + 1
    Case Else:                  NOth = NOth + 1
    End Select
    NTot = NTot + 1
Next
CmpCntDrzP = Array(P.Name, NTot, NMod, NCls, NDoc, NFrm, NOth)
End Function

Function Cmp(Cmpn) As VBComponent
Set Cmp = CPj.VBComponents(Cmpn)
End Function

Function NClsP%(): NClsP = NClszP(CPj): End Function
Function NCmpP%(): NCmpP = NCmpzP(CPj): End Function
Function NModP%(): NModP = NModzP(CPj): End Function
'---============================================

Function NCmpzP%(P As VBProject)
If P.Protection = vbext_pp_locked Then Exit Function
NCmpzP = P.VBComponents.Count
End Function

Function NModzP%(P As VBProject): NModzP = NCmpzTy(P, vbext_ct_StdModule): End Function
Function NClszP%(P As VBProject): NClszP = NCmpzTy(P, vbext_ct_ClassModule): End Function
Function NDoczP%(P As VBProject): NDoczP = NCmpzTy(P, vbext_ct_Document): End Function

Function NCmpzTy%(P As VBProject, Ty As vbext_ComponentType)
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
Dim O%
For Each C In P.VBComponents
    If C.Type = Ty Then O = O + 1
Next
NCmpzTy = O
End Function

Function NOthCmpzP%(P As VBProject)
NOthCmpzP = NCmpzP(P) - NClszP(P) - NModzP(P) - NDoczP(P)
End Function
