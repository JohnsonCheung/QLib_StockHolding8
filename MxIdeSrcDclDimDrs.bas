Attribute VB_Name = "MxIdeSrcDclDimDrs"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcDclDimDrs."
Public Const DimFF$ = "DimItm V Vsf"

Private Sub DimDrsP__Tst()
GoSub Z1
Exit Sub
Z1:
    Brw QSrt(AwDis(StrCol(DimDrsP, "Vsf")))
    Return
Z:  BrwDrs DimDrsP
    Return
End Sub

Function DimDrsP() As Drs
DimDrsP = DimDrszP(CPj)
End Function
Function DimDrszP(P As VBProject) As Drs
DimDrszP = DimDrs(DimItmAyzS(SrczP(P)))
End Function
Function DimDrs(DimItmAy$()) As Drs
DimDrs = DrszFF(DimFF, DimDy(DimItmAy))
End Function

Function DimDy(DimItmAy$()) As Variant()
Dim I: For Each I In Itr(DimItmAy)
    PushI DimDy, DimDr(I)
Next
End Function

Function DimDr(DimItm) As Variant()
With S12oDimnqVsfx(DimItm)
DimDr = Array(DimItm, .S1, .S2)
End With
End Function
