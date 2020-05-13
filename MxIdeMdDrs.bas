Attribute VB_Name = "MxIdeMdDrs"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdDrs."
 Public Const MdnFF$ = "Pjn Ns MdTy Mdn NMdLn"
Public Const Mdn9FF$ = MdnFF & " CLibv CNsv CModv IsCModEr"
  Public Const MdFF$ = Mdn9FF & " NMth NPub NPrv NFrd Mthnn"

Function MdnDrzM(M As CodeModule) As Variant()
MdnDrzM = MdnDr(PjnzM(M), _
    CNsv(Dcl(M)), _
    ShtMdTy(M), _
    Mdn(M), _
    M.CountOfLines)
End Function

Function MdnDr(Pjn$, Ns$, MdTy$, Mdn$, NMdLn&) As Variant()
MdnDr = Array(Pjn, Ns, MdTy, Mdn, NMdLn)
End Function
'--
Private Sub MdnDrsP__Tst():  BrwDrsN MdnDrsP:  End Sub
Private Sub Mdn9DrsP__Tst(): BrwDrsN Mdn9DrsP: End Sub
Private Sub MdDrsP__Tst():   BrwDrsN MdDrsP:   End Sub

Function MdnDrsP(Optional MdnPatn$) As Drs:   MdnDrsP = MdnDrszP(CPj, MdnPatn$): End Function
Function Mdn9DrsP(Optional MdnPatn$) As Drs: Mdn9DrsP = Mdn9Drs(CPj, MdnPatn):   End Function
Function MdDrsP(Optional MdnPatn$) As Drs:     MdDrsP = MdDrs(CPj, MdnPatn):     End Function

'--
Function MdnDrszP(P As VBProject, Optional MdnPatn$) As Drs
MdnDrszP = DrszFF(MdnFF, W1Dy(P))
End Function

Private Function W1Dy(P As VBProject, Optional MdnPatn$) As Variant()
Dim N: For Each N In AwPatn(Itn(P.VBComponents), MdnPatn)
    Push W1Dy, MdnDrzM(P.VBComponents(N).CodeModule)
Next
End Function
'--
Function MdDrs(P As VBProject, Optional MdnPatn$) As Drs
Dim D As Drs: D = Mdn9Drs(P, MdnPatn)
Dim Dy()
    Dim IxMdn%: IxMdn = IxzAy(D.Fny, "Mdn")
    Dim Dr: For Each Dr In Itr(D.Dy)
        Dim Mdn$: Mdn = Dr(IxMdn)
        Dim S$(): S = Src(PjMd(P, Mdn))
        Dim L$(): L = MthlnyzS(S)
        With MthCntzMthln(L, Si(S))
            Dim NMth%: NMth = .NPub + .NPrv + .NFrd
            PushI Dy, AddAy(Dr, Array(NMth, .NPub, .NPrv, .NFrd, Mthnn(S)))
        End With
    Next
MdDrs = AddColzFFDy(D, "NMth NPub NPrv NFrd Mthnn", Dy)
End Function

Function Mdn9Drs(P As VBProject, Optional MdnPatn$) As Drs
Const CSub$ = CMod & ".Mdn9Drs"
Dim ODy()
    Dim Drs As Drs: Drs = MdnDrszP(P, MdnPatn)
    Dim IxMdn%: IxMdn = EleIx(Drs.Fny, "Mdn")
    Dim Dr: For Each Dr In Drs.Dy
        Dim Mdn$: Mdn = Dr(IxMdn)
        Dim D$(): D = Dcl(P.VBComponents(Mdn).CodeModule)
        Dim ICModv$: ICModv = CModv(D)
        Dim IIsCModEr As Boolean: IIsCModEr = ICModv <> Mdn
        'Pjn Ns MdTy Mdn NMdLn
        'Pjn Ns MdTy Mdn NMdLn CLibv CNsv CModv IsCModEr
        PushI ODy, AddAyAp(Dr, CLibv(D), CNsv(D), ICModv, IIsCModEr)
    Next
Dim Bef As Drs: Bef = DrszFF(Mdn9FF, ODy)
Mdn9Drs = SrtDrs(Bef)
Insp CSub, "Before and After sort", "Bef Aft", FmtDrsR(Bef), FmtDrsR(Mdn9Drs)
End Function
