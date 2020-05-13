Attribute VB_Name = "MxIdeSrcLisMth"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Mth.Lis"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcLisMth."

Sub LisFunP(Patn$)
Dim A As Drs: A = FunDrsP
BrwDrs DwPatn(A, "Mthn", Patn)
End Sub

Sub LisPFunRetAs(RetAsPatn$)
':PFun: :Cml
Dim RetSfx As Drs: RetSfx = AddMthColRetTyn(PFunDrsP)
Dim Patn As Drs: Patn = DwPatn(RetSfx, "RetSfx", RetAsPatn)
Dim T50 As Drs: T50 = DwTopN(Patn)
BrwDrs T50
End Sub

Sub LisPPrpRetAs(RetAsPatn$)
Dim S As Drs: S = PFunDrsP
Dim RetSfx As Drs: RetSfx = AddMthColRetTyn(S)
Dim Pub As Drs: Pub = DwEqExl(RetSfx, "Mdy", "Pub")
Dim Fun As Drs: Fun = DwEqExl(Pub, "Ty", "Get")
Dim Patn As Drs: Patn = DwPatn(Fun, "RetSfx", RetAsPatn)
Dim T50 As Drs: T50 = DwTopN(Patn)
BrwDrs T50
End Sub

Sub BrwMthP(W As WhMth)
LisMth CPj, W, eBrwOup, Top:=0
End Sub

Sub VcMthP(W As WhMth)
LisMth CPj, W, OupTy:=eVcOup, Top:=0
End Sub

Sub DmpMthP(W As WhMth, Optional Top% = 50)
LisMth CPj, W, eDmpOup, Top
End Sub

Private Sub LisMth(P As VBProject, W As WhMth, OupTy As eOupTy, Top%)
Dim D As Drs
    D = MthLisWh(MthLisDrszP(CPj), W)
    D = DwTopN(D, Top)
Dim O$()
    O = FmtDrsR(D)
OupAy O, OupTy
End Sub

Sub VcAllMthP()
VcMthP WhAllMth
End Sub

Sub BrwAllMthP()
BrwMthP WhAllMth
End Sub
