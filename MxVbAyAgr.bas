Attribute VB_Name = "MxVbAyAgr"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CNs$ = "Ay.Agr"
Const CMod$ = CLib & "MxVbAyAgr."

Sub BrwMdLnCntAgrP()
End Sub
Function MdLnCntAgrP() As Dictionary
Set MdLnCntAgrP = DiAgrzVal(MdLnCntP)
End Function

Function MdLnCntP() As Long()
MdLnCntP = MdLnCntAyzP(CPj)
End Function

Function MdLnCntAyzP(P As VBProject) As Long()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI MdLnCntAyzP, C.CodeModule.CountOfLines
Next
End Function

Function CntNo0&(NumAy)
Dim O&
Dim V: For Each V In Itr(NumAy)
    If V <> 0 Then O = O + 1
Next
CntNo0 = O
End Function

Function DiAgrzVal(NumAy) As Dictionary
'Ret : Agr Val ! where *Arg has Cnt Avg Max Min Sum
Dim O As New Dictionary
Dim Sum#: Sum = AySum(NumAy)
Dim NNo0&: NNo0 = CntNo0(NumAy)
Dim N&: N = Si(NumAy)
Dim AvgAll#, AvgNo0#
If N <> 0 Then AvgAll = Sum / N
If NNo0 <> 0 Then AvgNo0 = Sum / NNo0

O.Add "CntNo0", NNo0
O.Add "CntAll", N
O.Add "AvgNo0", AvgNo0
O.Add "AvgAll", AvgAll
O.Add "Sum", Sum
O.Add "Max", MaxEle(NumAy)
O.Add "Min", MinEle(NumAy)
O.Add "MinGT0", MinElezGT0(NumAy)
Set DiAgrzVal = O
End Function
