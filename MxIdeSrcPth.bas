Attribute VB_Name = "MxIdeSrcPth"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcPth."

Function SrcPthzCmp$(A As VBComponent)
SrcPthzCmp = SrcPthzP(PjzCmp(A))
End Function

Function SrcPjf$(SrcPth$)
ChkIsSrcPth SrcPth, "SrcPjf"
SrcPjf = RmvFstChr(Fdr(ParPth(ParPth(SrcPth))))
End Function

Function SrcPth$(Pjf$)
SrcPth = EnsPth(AssPth(Pjf) & ".src")
End Function

Sub EnsSrcPthzP(P As VBProject)
EnsAllFdr SrczP(P)
End Sub

Function SrcPthzDistPj$(DistPj As VBProject)
Dim P$: P = Pjp(DistPj)
SrcPthzDistPj = AddFdrAp(UpPth(P, 1), ".Src", Fdr(P))
End Function

Sub BrwSrcPthP()
BrwPth SrcPthP
End Sub

Function SrcPthP$()
SrcPthP = SrcPthzP(CPj)
End Function

Function SrcPthzAcs$(A As Access.Application)
SrcPthzAcs = SrcPthzP(MainPj(A))
End Function

Function SrcPthzP$(P As VBProject)
SrcPthzP = SrcPth(Pjf(P))
End Function

Function IsSrcPth(Pth) As Boolean
Dim F$: F = Fdr(Pth)
If Not HasExtSS(F, ".xlam .accdb") Then Exit Function
IsSrcPth = Fdr(ParPth(Pth)) = ".Src"
End Function
