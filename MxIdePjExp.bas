Attribute VB_Name = "MxIdePjExp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdePjExp."

Sub ExpRf(P As VBProject)
WrtAy RfSrc(P), FrfzP(P)
End Sub

Sub ExpPjf(Pjf, Optional Xls As Excel.Application, Optional Acs As Access.Application)
Stamp "ExpPj: Begin"
Stamp "ExpPj: Pjf " & Pjf
Select Case True
Case IsFxa(Pjf): ExpToFx Pjf
Case IsFba(Pjf): ExpToFb Pjf
End Select
Stamp "ExpPj: End"
End Sub

Sub ExpToFb(Fb): ExpAcs AcszFb(Fb): End Sub
Sub ExpToFx(Fx): ExpXls NwXlszFx(Fx): End Sub
Sub ExpPjP(): ExpPj CPj: End Sub
Sub ExpPj(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    DoEvents
    ExpMd C.CodeModule
Next
End Sub
Sub ExpMd(M As CodeModule): M.Parent.Export SrcFfnzM(M): End Sub

Sub ExpXls(X As Excel.Application)

End Sub
Sub ExpCAcs(): ExpAcs Acs: End Sub

Sub ExpAcs(A As Access.Application, Optional ToFba$)
Const CSub$ = CMod & "ExpPjzP"
Dim Fb$: Fb = A.CurrentDb.Name
Dim Pj As VBProject: Set Pj = MainPj(A)
Dim P$: P = SrcPthzP(Pj)
CpyFfnToPth Fb, EnsPth(P), OvrWrt:=True
InfLn CSub, "... Clr src pth":       EnsAllFdr P
                                     DltAllPthFil P
InfLn CSub, "... Cpy pj to src pth": CpyFfnToPth Pj.FileName, P
InfLn CSub, "... Exp src":           ExpPj Pj
InfLn CSub, "... Exp rf":            ExpRf Pj
InfLn CSub, "... Exp frm":           ExpAllFrm A, P
InfLn CSub, "... Exp rpt":           ExpAllRpt A, P
InfLn CSub, "Done"
End Sub
