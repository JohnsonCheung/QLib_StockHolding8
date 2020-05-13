Attribute VB_Name = "MxNDclLin"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxNDclLin."
Public Const NDclLinFF$ = "Mdn NDclLin"

Function NDclLinDrsP() As Drs
NDclLinDrsP = NDclLinDrszP(CPj)
End Function

Function NDclLinDrszP(P As VBProject) As Drs
NDclLinDrszP = DrszFF(NDclLinFF, NDclLinDy(P))
End Function

Function NDclLinDy(P As VBProject) As Variant()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI NDclLinDy, Array(C.Name, NDclLin(C.CodeModule))
Next
End Function

Function NDclLinzS%(Src$())
'Assume FstMth cannot have Mrmk
Dim O&: O = FstMthix(Src)
NDclLinzS = O - NLinNonCdAbovezS(Src, O)
End Function

Function NDclLin%(M As CodeModule) 'Assume FstMth cannot have Mrmk
NDclLin = M.CountOfDeclarationLines
End Function

Function NLinNonCdAbove&(M As CodeModule, Lno&)
Dim O%
Dim J&: For J = Lno To 1 Step -1
    If Not IsCdLn(M.Lines(J, 1)) Then NLinNonCdAbove = O: Exit Function
    O = O + 1
Next
NLinNonCdAbove = Lno
End Function
Function NLinNonCdAbovezS&(Src$(), Ix&)
Dim O%
Dim J&: For J = Ix To 0 Step -1
    If Not IsCdLn(Src(J)) Then NLinNonCdAbovezS = O: Exit Function
    O = O + 1
Next
NLinNonCdAbovezS = Ix
End Function
