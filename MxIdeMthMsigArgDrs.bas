Attribute VB_Name = "MxIdeMthMsigArgDrs"
Option Explicit
Option Compare Text

Public Const ArgFF$ = "Mthn No Nm IsOpt IsByVal IsPmAy IsAy TyChr AsTy DftVal"

Function ArgDrszMthL(Mthlny$()) As Drs
ArgDrszMthL = DrszFF(ArgFF, ArgDyzMthL(Mthlny))
End Function

Function ArgDrszM(M As CodeModule) As Drs
Dim L$(): L = StrCol(MthDrszM(M), "Mthln")
     ArgDrszM = DrszFF(ArgFF, ArgDyzMthL(L))
End Function

Function ArgDy(Mthln) As Variant()
Dim Pm$: Pm = BetBkt(Mthln)
Dim A$(): A = AmTrim(SplitCommaSpc(Pm))
Dim Mthn$: Mthn = MthnzL(Mthln)
Dim Arg$, Dy(), ArgNo%: For ArgNo = 1 To Si(A)
    Arg = A(ArgNo - 1)
    PushI ArgDy, ArgDr(Arg, ArgNo, Mthn)
Next
End Function

Function ArgDrszP(P As VBProject) As Drs
ArgDrszP = DrszFF(ArgFF, ArgDyzMthL(MthlnyzP(P)))
End Function

Private Sub ArgDrsP__Tst()
BrwDrs ArgDrsP
End Sub

Function ArgDrsP() As Drs
ArgDrsP = ArgDrszP(CPj)
End Function

Function ArgDrs(Mthln) As Drs
ArgDrs = DrszFF("Mthn Nm ArgNo IsOpt IsByVal IsPmAy IsAy TyChr RetSfx DftVal", ArgDy(Mthln))
End Function

Function ArgDr(ArgStr$, ArgNo%, Mthn$) As Variant()
Dim A As Arg: A = ArgzS(ArgStr)
Dim M As eArgm: M = A.Argm
Dim IsOpt   As Boolean:   IsOpt = M = eOptByRefArgm Or M = eOptByValArgm
Dim IsByVal As Boolean: IsByVal = M = eByValArgm Or M = eOptByValArgm
Dim IsPmAy  As Boolean:  IsPmAy = M = ePmArgm
Dim Nm$:                     Nm = A.Argn
Dim TyChr$:               TyChr = A.Ty.TyChr
Dim IsAy    As Boolean:    IsAy = A.Ty.IsAy
Dim Tyn$:                   Tyn = A.Ty.Tyn
ArgDr = Array(Mthn, ArgNo, Nm, IsOpt, IsByVal, IsPmAy, IsAy, TyChr, Tyn, A.Dft)
End Function

Function ArgDrsV() As Drs
ArgDrsV = ArgDrszMthL(MthlnyV)
End Function

Function ArgDyzMthL(Mthlny$()) As Variant()
Dim L: For Each L In Itr(Mthlny)
    PushIAy ArgDyzMthL, ArgDy(L)
Next
End Function
