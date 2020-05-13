Attribute VB_Name = "MxIdeSrcEndLn"
Option Compare Text
Option Explicit
Const CNs$ = "Src.Itm"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcEndLn."
Function SrcEix&(Src$(), SrcItmBix)
If SrcItmBix < 0 Then SrcEix = -1: Exit Function
Const CSub$ = CMod & "MthEix"
':MthEix: :Ix ! #Mth-End-Ix# it is an @Src-Ix pointing to the Las-Ln of the mth pointed by @SrcItmBix
Dim Endln$: Endln = SrcEndlnzS(Src, SrcItmBix)
If HasSubStr(Src(SrcItmBix), Endln) Then SrcEix = SrcItmBix: Exit Function
Dim O&: For O = SrcItmBix + 1 To UB(Src)
   If HasPfx(Src(O), Endln) Then SrcEix = O: Exit Function
Next
Thw CSub, "Cannot find VbMthEndLin", "MthEndLin SrcItmBix Src", Endln, SrcItmBix, Src
End Function

Function SrcEno&(M As CodeModule, SrcItmLno)
Const CSub$ = CMod & "EndLnozM"
Dim Endln$, O&
Endln = SrcEndlnzM(M, SrcItmLno)
If HasSubStr(M.Lines(SrcItmLno, 1), Endln) Then SrcEno = SrcItmLno: Exit Function
For O = SrcItmLno + 1 To M.CountOfLines
   If HasPfx(M.Lines(O, 1), Endln) Then SrcEno = O: Exit Function
Next
Thw CSub, "Cannot find EndLno", "MthEndLin @SrcItmLno @Mdn", Endln, SrcItmLno, Mdn(M)
End Function

Function SrcEndlnzS$(Src$(), B)
SrcEndlnzS = SrcEndln(SrcItm(Src(B)))
End Function

Function SrcEndlnzM$(M As CodeModule, SrcItmLno)
SrcEndlnzM = SrcEndln(SrcItm(M.Lines(SrcItmLno, 1)))
End Function

Function SrcEndln$(SrcItm$)
ChkIsVbItm SrcItm, "SrcEndln"
SrcEndln = "End " & SrcItm
End Function
