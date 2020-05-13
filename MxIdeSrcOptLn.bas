Attribute VB_Name = "MxIdeSrcOptLn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcOptLn."
Const ExpOptLin$ = "Option Explicit"
Const CmpBinOptLin$ = "Option Compare Binary"
Const CmpDbOptLin$ = "Option Compare Database"
Const CmpTxtOptLin$ = "Option Compare Text"

Sub Ens3OptM()
Ens3OptzM CMd
End Sub

Sub Ens3Opt()
Ens3OptzP CPj
End Sub

Sub Ens3OptzP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    Ens3OptzM C.CodeModule
Next
End Sub

Private Sub Ens3OptzM__Tst()
Dim M As CodeModule
Const Mdn$ = "AA"
GoSub Setup
GoSub T0
GoSub Clean
Exit Sub
T0:
    Set M = Md(Mdn)
    GoTo Tst
Tst:
    Ens3OptzM M
    Return
Setup:
    AddClsnn Mdn
    Return
Clean:
    RmvMd Mdn
    Return
End Sub

Sub Ens3OptzM(M As CodeModule)
If IsEmpMd(M) Then Exit Sub
DltOptLinzM M, CmpDbOptLin
DltOptLinzM M, CmpBinOptLin
EnsOptLinzM M, CmpTxtOptLin
EnsOptLinzM M, ExpOptLin
End Sub

Private Sub AftOptqImplLno__Tst()
Dim M As CodeModule
GoSub T0
Exit Sub
T0:
    Set M = Md("ATaxExpCmp_OupTblGenr")
    Ept = 2&
    GoTo Tst
Tst:
    Act = AftOptqImplLno(M)
    C
    Return
End Sub

Function IxoAftOptqImplzS&(Src$())
Dim Fnd As Boolean, J%, IsOpt As Boolean, L$
For J = 0 To UB(Src)
    L = Src(J)
    'IsOpt = IsLn_OfOpt_OrImpl_OrBlnk(L)
    Select Case True
    Case Fnd And IsOpt:
    Case Fnd: IxoAftOptqImplzS = J: Exit Function
    Case IsOpt: Fnd = True
    End Select
Next
IxoAftOptqImplzS = J
End Function

Function IsOptOrImplLin(Ln) As Boolean
Select Case True
Case IsOptln(Ln), IsImpln(Ln): IsOptOrImplLin = True
End Select
End Function

Function AftOptqImplLno&(M As CodeModule)
Dim N%: N = M.CountOfDeclarationLines
Dim J%: For J = 1 To N
    Dim L$: L = M.Lines(J, 1)
    If Not IsOptOrImplLin(L) Then AftOptqImplLno = J: Exit Function
Next
AftOptqImplLno = N + 1
End Function

Function OptLno%(M As CodeModule, OptLin)
Dim J&
For J = 1 To M.CountOfDeclarationLines
   If M.Lines(J, 1) = OptLin Then OptLno = J: Exit Function
Next
End Function

Sub EnsOptLinzM(M As CodeModule, OptLin)
Const CSub$ = CMod & "EnsOptLin"
If M.CountOfLines = 0 Then Exit Sub
If OptLno(M, OptLin) > 0 Then Exit Sub
M.InsertLines 1, OptLin
InfLn CSub, "[" & OptLin & "] is Inserted", "Md", Mdn(M)
End Sub

Sub DltOptLinzM(M As CodeModule, OptLin)
Const CSub$ = CMod & "DltOptLin"
Dim I%: I = OptLno(M, OptLin)
If I = 0 Then Exit Sub
M.DeleteLines I
Inf CSub, "[" & OptLin & "] line is deleted", "Md Lno", Mdn(M), I
End Sub
