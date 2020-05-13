Attribute VB_Name = "MxIdeSrcCnstStrv"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Dcl.3Cnst"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcCnstStrv."

Function StrCnstvzL$(Ln, Cnstn$)
If IsCnstNmLin(Ln, Cnstn) Then StrCnstvzL = BetDblQ(Ln)
End Function
Function StrCnstvzM(M As CodeModule, Cnstn$)
Dim L, O$, J%: For J = 1 To M.CountOfLines
    O = StrCnstvzL(M.Lines(J, 1), Cnstn)
    If O <> "" Then StrCnstvzM = O: Exit Function
Next
End Function

Function StrCnstv(Dcl$(), Cnstn$)
Dim L, O$: For Each L In Itr(Dcl)
    O = StrCnstvzL(L, Cnstn)
    If O <> "" Then StrCnstv = O: Exit Function
Next
End Function

Function StrCnstvzLin$(Ln)
If IsStrCnstLin(Ln) Then StrCnstvzLin = BetDblQ(Ln)
End Function

Sub StrCnstv__Tst()
Dim O$()
Dim L: For Each L In SrczP(CPj)
    PushNB O, StrCnstvzLin(L)
Next
BrwAy O
End Sub
