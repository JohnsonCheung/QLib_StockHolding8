Attribute VB_Name = "MxIdeMthCntDrs"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthCntDrs."
Const CntgMthPP$ = "NPSub NPFun NPPrp NPrvSub NPrvFun NPrvPrp NFrdSub NFrdFun NFrdPrp"
Public Const MthCntFF$ = "Lib Mdn NLn NMth NPSub NPFun NPPrp NPrvSub NPrvFun NPrvPrp NFrdSub NFrdFun NFrdPrp"
Type CntgMth
    Lib As String
    Mdn As String
    NPSub As Integer
    NPFun As Integer
    NPPrp As Integer
    NPrvSub As Integer
    NPrvFun As Integer
    NPrvPrp As Integer
    NFrdSub As Integer
    NFrdFun As Integer
    NFrdPrp As Integer
End Type

Function MthCntDrsP(Optional MdnPatn$ = ".+", Optional SrtCol$ = "Mdn") As Drs
MthCntDrsP = MthCntDrszP(CPj, MdnPatn, SrtCol)
End Function

Function MthCntDrszP(P As VBProject, MdnPatn$, SrtCol$) As Drs
Dim R As RegExp: Set R = Rx(MdnPatn, IgnoreCase:=True)
Dim C As VBComponent, Dy(): For Each C In P.VBComponents
    If R.Test(C.Name) Then
        PushI Dy, MthCntDr(C.CodeModule)
    End If
Next
Dim D As Drs: D = DrszFF(MthCntFF, Dy)
MthCntDrszP = SrtDrs(D, SrtCol)
End Function

Function NMth%(A As CntgMth)
With A
NMth = .NPSub + .NPFun + .NPPrp + .NPrvSub + .NPrvFun + .NPrvPrp + .NFrdSub + .NFrdFun + .NFrdPrp
End With
End Function
Function FmtCntgMth(A As CntgMth, Optional Hdr As eAnyHdr)
Dim Pfx$: If Hdr = eHdrNo Then Pfx = "Pub* | Prv* | Frd* : *{Sub Fun Frd} "
With A
Dim N%: N = NMth(A)
FmtCntgMth = JnAp(" | ", Pfx, .Mdn, N) & " | " & JnSpcAp(.NPSub, .NPFun, .NPPrp, .NPrvSub, .NPrvFun, .NPrvPrp, .NFrdSub, .NFrdFun, .NFrdPrp)
End With
End Function

Function NMthzS%(Src$())
NMthzS = Si(Mthixy(Src))
End Function

Function NMthP%()
NMthP = NMthzP(CPj)
End Function

Function NMthM%()
NMthM = NMthzM(CMd)
End Function

Function NMthzP%(Pj As VBProject)
Dim O%, C As VBComponent
For Each C In Pj.VBComponents
    O = O + NMthzS(Src(C.CodeModule))
Next
NMthzP = O
End Function

Function MthCntDr(M As CodeModule) As Variant()
Const CSub$ = CMod & "MthCntDr"
Dim S$(): S = Src(M)
Dim Mth$(): Mth = Mthlny(S)
Dim L: For Each L In Itr(Mth)
    With Mthn3zL(L)
        Dim Prv As Boolean: Prv = False
        Dim Pub As Boolean: Pub = False
        Dim Frd As Boolean: Frd = False
        Dim Fun As Boolean: Fun = False
        Dim Sbr As Boolean: Sbr = False
        Dim Prp As Boolean: Prp = False
        Select Case .ShtMdy
        Case "Prv": Prv = True
        Case "Pub": Pub = True
        Case "Frd": Frd = True
        Case Else: Thw CSub, "Out of valid value: Prv PUb Frd", "ShtMdy", .ShtMdy
        End Select
        Select Case ShtMthKdzShtMthTy(.ShtTy)
        Case "Fun": Fun = True
        Case "Sub": Sbr = True
        Case "Prp": Prp = True
        Case Else: Thw CSub, "Out of valid value: Sub Fun Prp", "ShtMdy", .ShtMdy
        End Select
    End With
            
    Select Case True
        Case Pub And Sbr: Dim NPSub%: NPSub = NPSub + 1
        Case Pub And Fun: Dim NPFun%: NPFun = NPFun + 1
        Case Pub And Prp: Dim NPPrp%: NPPrp = NPPrp + 1
        Case Prv And Sbr: Dim NPrvSub%: NPrvSub = NPrvSub + 1
        Case Prv And Fun: Dim NPrvFun%: NPrvFun = NPrvFun + 1
        Case Prv And Prp: Dim NPrvPrp%: NPrvPrp = NPrvPrp + 1
        Case Frd And Sbr: Dim NFrdSub%: NFrdSub = NFrdSub + 1
        Case Frd And Fun: Dim NFrdFun%: NFrdFun = NFrdFun + 1
        Case Frd And Prp: Dim NFrdPrp%: NFrdPrp = NFrdPrp + 1
        Case Else: Thw CSub, "Invalid Mthn3", "Mthln", L
    End Select
    Dim NMth%: NMth = NMth + 1
    If NPSub + NPFun + NPPrp + NPrvSub + NPrvFun + NPrvPrp + NFrdSub + NFrdFun + NFrdPrp <> NMth Then Stop
Next
Dim NLin&: NLin = Si(S)
Dim Mdn$: Mdn = MdnzM(M)
Dim Lib$: Lib = Bef(Mdn, "_")
MthCntDr = Array(Lib, Mdn, NLin, NMth, NPSub, NPFun, NPPrp, NPrvSub, NPrvFun, NPrvPrp, NFrdSub, NFrdFun, NFrdPrp)
End Function

Sub LisMdP(Optional MdnPatn$ = ".+", Optional SrtCol$ = "Mdn")
LisMdzP CPj, MdnPatn, SrtCol
End Sub

Sub LisMdzM(M As CodeModule)
DmpDrs MthCntDrszM(M)
End Sub

Function MthCntDrszM(M As CodeModule) As Drs
MthCntDrszM = DrszFF(MthCntFF, Av(MthCntDr(M)))
End Function


Sub LisMdM()
LisMdzM CMd
End Sub

Sub LisMdzP(P As VBProject, MdnPatn$, SrtCol$, Optional OupTy As eOupTy = eOupTy.eBrwOup)
LisAy FmtDrsR(MthCntDrszP(P, MdnPatn, SrtCol)), OupOpt("LisMd_", OupTy)
End Sub

Function NMthzM%(M As CodeModule)
NMthzM = NMthzS(Src(M))
End Function

Function NSrcLinPj&(P As VBProject)
Dim O&, C As VBComponent
For Each C In P.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
NSrcLinPj = O
End Function

Function NPubMthzS%(Src$())
NPubMthzS = NItr(PubMthlnItr(Src))
End Function

Function NPubMthzM%(M As CodeModule)
NPubMthzM = NPubMthzS(Src(M))
End Function

Function NPubMthzV%(A As Vbe)
Dim O%, P As VBProject
For Each P In A.VBProjects
    O = O + NPubMthzP(P)
Next
NPubMthzV = O
End Function

Property Get NPubMthV%()
NPubMthV = NPubMthzV(CVbe)
End Property

Function NPubMthzP%(P As VBProject)
Dim O%, C As VBComponent
For Each C In P.VBComponents
    O = O + NPubMthzM(C.CodeModule)
Next
NPubMthzP = O
End Function
