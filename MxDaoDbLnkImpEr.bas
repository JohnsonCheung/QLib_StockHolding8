Attribute VB_Name = "MxDaoDbLnkImpEr"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDbLnkEr."
#If Doc Then
'Cml
' Vinp #V-Lvl-Fmt-Er#
' Eu   #Error-Udt# Use as Sfx-Cml showing that is it a Udt
' Et   #Error-Ty#  Use as Sfx-Enum showing that it is a Enum for Error
' Bexp #Bool-Express# a string can evaluated to boolean
' Er   #Error#     Use as Sfx-Varn showing that is an error of type :Ly
#End If
'-- Er
Private Enum eFxwEt: eWsNFnd: End Enum
Private Enum eFbtEt: eFilNFnd: End Enum
Private Enum eInpEt: eFilNFnd: End Enum
'-- FbEu
Private Type FbnDupEu: Lix As Integer: End Type
Private Type FbnMisEu: Lix As Integer: End Type
Private Type FbTblDupEu: Lix As Integer: End Type
Private Type FbTblMisEu: Lix As Integer: End Type
Private Type FbStruMisEu: Lix As Integer: End Type
Private Type FbtEu
    FbnMis() As FbnMisEu
    FbnDup() As FbnDupEu
    FbTblDup() As FbTblDupEu
    FbTblMis() As FbTblMisEu
    StruMis() As FbStruMisEu
End Type
'-- FxwEu
Private Type FxTblDupEu: Lix As Integer: End Type
Private Type FxnDupEu: Lix As Integer: End Type
Private Type FxnMisEu: Lix As Integer: End Type
Private Type FxStruMisEu: Lix As Integer: End Type
Private Type FxwEu
    FxTblDup As FxTblDupEu
End Type
'-- InpEu
Private Type InpnDupEu: Lix As Integer: End Type
Private Type InpfDupEu: Lix As Integer: End Type
Private Type InpfMisEu: Lix As Integer: End Type
Private Type InpfKdDupEu: Lix As Integer: End Type
Private Type InpEu
    InpnDup() As InpnDupEu
    InpfDup() As InpfDupEu
    InpfMis() As InpfMisEu
End Type
'-- StruEu
Private Type StruDupEu: Lix As Integer: End Type
Private Type StruMisEu: Lix As Integer: End Type
Private Type StruExaEu: Lix As Integer: End Type
Private Type StruFldDupEu: Lix As Integer: End Type
Private Type StruTyErEu: Lix As Integer: End Type
Private Type StruEu
    StruDup As StruDupEu
    StruMis() As StruMisEu
    StruExa() As StruExaEu
    FldDup() As StruFldDupEu
    TyEr() As StruTyErEu
End Type
Private Type BexpEu: Lix As Integer: End Type
Private Type OthEu:
    NoFxAndNoFb As Boolean
End Type
Private Type Er: Inp As InpEu: Fxw As FxwEu: Fbt As FbtEu: Stru As StruEu: Bexp As BexpEu: Oth As OthEu: End Type
Function LnkErzS(LnkImpSrc1$()) As String()

End Function

Function LnkEr(Src As Lnkis) As String()
'Fm: *InpFilSrc::SSAy{FilKd Ffn}
'Fm: *LnkImpSrc::IndLy{
'                  TblFx: {Fxt} [{Fxn}[.{Wsn}]] [{Stru}]
'                  TblFb: {Fbn} {Fbtt}
'                  Stru.{Stru}: {F} [{Ty}] [{Extn}]
'                  Tbl.Where: {T} {Bexp}
'                } @@
'-----------------------------------------------------------------------------------------------------------------------
Dim E As Er: E = UUEr(Src)
With E
Dim I$(): I = UUInpEr(.Inp)
Dim X$(): X = UUFxwEr(.Fxw)
Dim B$(): B = UUFbtEr(.Fbt)
Dim S$(): S = UUStruEr(.Stru)
Dim W$(): W = UUBexpEr(.Bexp)
Dim O$(): O = UUOthEr(.Oth)
End With
LnkEr = AddSyAp(I, X, B, S, W, O)
End Function
Private Function UUOthEu(S As Lnkis) As OthEu
With UUOthEu
    .NoFxAndNoFb = VothNoFxAndNoFb()
'    .NoFxAndNoFb = W1NoFxAndNoFb(Ipx, Ipb)
End With
End Function
Private Function VothNoFxAndNoFb() As Boolean
'If Si(Ipx.Dy) > 0 Then Exit Function
'If Si(Ipb.Dy) > 0 Then Exit Function
VothNoFxAndNoFb = True
End Function

Private Function UUOthEr(E As OthEu) As String()
PushIAy UUOthEr, W1NoFxAndNoFb(E.NoFxAndNoFb)
PushIAy UUOthEr, W1HdrEr(E)
End Function

Private Function W1NoFxAndNoFb(NoFxAndNoFb As Boolean) As String()
If Not NoFxAndNoFb Then Exit Function
PushI W1NoFxAndNoFb, ""
End Function

Private Function W1HdrEr(E As OthEu) As String()

End Function

Private Function UUEr(S As Lnkis) As Er
With UUEr
    .Bexp = UUBexpEu
    .Fbt = UUFbtEu(S)
    .Fxw = UUFxwEu
    .Inp = UUInpEu
    .Oth = UUOthEu(S)
    .Stru = UUStruEu
End With
End Function
Private Function UUInpEu() As InpEu

End Function

Private Function UUFxwEu() As FxwEu

End Function
Private Function UUBexpEu() As BexpEu

End Function

Private Function UUInpEr(E As InpEu) As String()
Dim A$(): A = W2FilKdDupEr()
Dim B$(): B = W2FfnDupEr()
Dim C$(): C = W2FfnMisEr()
UUInpEr = Sy(A, B, C)
End Function

Private Function W2FilKdDupEr() As String()
End Function
Private Function W2FfnDupEr() As String()
End Function
Private Function W2FfnMisEr() As String()
End Function
Private Function UUFbtEr(Eu As FbtEu) As String()

End Function
Private Function UUFxwEr(Eu As FxwEu) As String()
Dim A$(): A = W3TblDupEr()
Dim B$(): B = W3FxnDupEr()
Dim C$(): C = W3FxnMisEr()
Dim D$(): D = W3WsMisEr()
Dim E$(): E = W3WsMisFldEr()
Dim F$(): F = W3WsMisFldTyEr()
Dim G$(): G = W3StruMisEr()
UUFxwEr = AddSyAp(A, B, C, D, E, F, G)
End Function

Private Function W3TblDupEr() As String()

End Function
Private Function W3FxnDupEr() As String()

End Function
Private Function W3FxnMisEr() As String()

End Function
Private Function W3WsMisEr() As String()

End Function
Private Function W3WsMisFldEr() As String()

End Function
Private Function W3WsMisFldTyEr() As String()

End Function
Private Function W3StruMisEr() As String()

End Function
Private Function UUFbtEu(S As Lnkis) As FbtEu
With UUFbtEu
    .FbnDup = VfbnDupEu
    .FbnMis = VfbnMisEu
    .FbTblDup = VfbTblDupEu
    .FbTblMis = VfbTblMisEu
    .StruMis = VfbStruMisEu
End With
End Function
Private Function VfbStruMisEu() As FbStruMisEu()

End Function
Private Function VfbTblMisEu() As FbTblMisEu()

End Function
Private Function VfbTblDupEu() As FbTblDupEu()

End Function
Private Function VfbnMisEu() As FbnMisEu()

End Function
Private Function UUStruEu() As StruEu
With UUStruEu
    .StruDup = VstruDupEu ' IpsHdStru
    .StruMis = VstruMisEu
    .StruExa = VstruExaEu
    .FldDup = VstruFldDupEu
    .TyEr = VstruTyErEu
End With
End Function
Private Function VstruDupEu() As StruDupEu

End Function
Private Function UUStruEr(Eu As StruEu) As String()
With Eu
Dim A$(): A = W4DupEr()
Dim B$(): B = W4MisEr
Dim C$(): C = W4ExaEr
Dim D$(): D = W4NoFldEr
Dim E$(): E = W4FldDupEr
Dim F$(): F = W4TyEr
End With
UUStruEr = AddSyAp(A, B, C, D, E, F)
End Function
Private Function W4DupEr() As String()
End Function
Private Function W4MisEr() As String()
End Function
Private Function W4ExaEr() As String()
End Function
Private Function W4NoFldEr() As String()
End Function
Private Function W4FldDupEr() As String()
End Function
Private Function W4TyEr() As String()
End Function
Private Function UUBexpEr(E As BexpEu) As String()
Dim A$(): A = W5TblDupEr()
Dim B$(): B = W5TblExaEr()               ' tbl.wh is more
Dim C$(): C = W5EmpEr()                   ' with tbl nm but no Bexp
UUBexpEr = AddSyAp(A, B, C)
End Function

Private Function W5TblDupEr() As String()

End Function
Private Function W5TblExaEr() As String()

End Function
Private Function W5EmpEr() As String()

End Function
'---===============================================

Private Function VfxWsMisFld(IpxfMis As Drs, ActWsf As Drs) As String()
If NoReczDrs(IpxfMis) Then Exit Function
Dim OFx$(), OFxn$(), OWs$(), O$(), Fxn, Fx$, Ws$, Mis As Drs, Act As Drs, J%, O1$()
AsgCol IpxfMis, "Fxn Fx Ws", OFxn, OFx, OWs
'---=
PushI O, "Some columns in ws is missing"
For Each Fxn In OFxn
    Fxn = OFxn(J)
    Fx = OFx(J)
    Ws = OWs(J)
    Mis = Dw3EqE(IpxfMis, "Fxn Fx Ws", Fxn, Fx, Ws)
    Act = Dw3EqE(ActWsf, "Fxn Fx Ws", Fxn, Fx, Ws)
    '-
    
    X "Fxn    : " & Fxn
    X "Fx pth : " & Pth(Fx)
    X "Fx fn  : " & Fn(Fx)
    X "Ws     : " & Ws
    X NmvzDrs("Mis col: ", Mis)
    X NmvzDrs("Act col: ", Act)
    'PushIAy O, AmTab(XX)
    J = J + 1
Next
VfxWsMisFld = O
'Insp "QDao_Lnk_LnkEr.LnkEr", "Inspect", "Oup(VfxWsMisFld) ExWsMisFld IpxfMis ActWsf",ExWsMisFld, ExWsMisFld, FmtDrs(IpxfMis), FmtDrs(ActWsf): Stop
End Function
Private Function VfxWsMisFldTy(Ipxf As Drs, ActWsf As Drs) As String()
'Fm IpxFld : Fxn Ws Stru Ipxf Ty Fx ! Where HasFx and HasWs and Not HasFld
'Fm WsActf : Fxn Ws Ipxf Ty @@
'Dim OFxn$(), J%, Fxn$, Fx$, Act$(), Lno&(), Ws$(), ActWsf()
'OFxn = AwDis(StrCol(IpXB, "Fxn"))
''---=
'If Si(OFxn) = 0 Then Exit Function
'PushI VfxWsMis, "Some expected ws not found"
'For J = 0 To UB(OFxn)
'    Fxn = OFxn(J)
'    Fx = ValzDrs(IpXB, "Fxn", Fxn, "Fx")
'    ActWsf = DwEqSel(IpXB, "Fxn", Fxn, "L Ws").Dy
'    Lno = LngAyzDyC(ActWsf, 0)
'    Ws = SyzDyC(ActWsf, 1)
'
'    Act = AmzRmvT1(AwT1(WsAct, Fxn)) '*WsActPerFxn::Sy{WsAct}
'    PushIAy VfxWsMis, XMisWs_OneFx(Fxn, Fx, Lno, Ws, Act)
'Next
'Insp "QDao_Lnk_LnkEr.LnkEr", "Inspect", "Oup(VfxWsMisFldTy) ExWsMisFldTy Ipxf ActWsf",ExWsMisFldTy, ExWsMisFldTy, FmtDrs(Ipxf), FmtDrs(ActWsf): Stop
End Function

Private Function VfxWsMis(IpxMis As Drs, ActWs As Drs) As String()
'@ActWs : Fxn Ws @@
Dim OFxn$(), J%, Fxn$, Fx$, Act$(), Lno&(), Ws$(), ActWsnn$, IpxMisi As Drs, O$()
OFxn = AwDis(StrCol(IpxMis, "Fxn"))
'---=
If Si(OFxn) = 0 Then Exit Function
PushI O, "Some expected ws not found"
For J = 0 To UB(OFxn)
    Fxn = OFxn(J)
    Fx = ValzDrs(IpxMis, "Fxn", Fxn, "Fx")
    IpxMisi = DwEqSel(IpxMis, "Fxn", Fxn, "L Ws")
    ActWsnn = Termln(FstCol(DwEqExl(ActWs, "Fxn", Fxn)))
    '-
    X "Fxn    : " & Fxn
    X "Fx pth : " & Pth(Fx)
    X "Fx fn  : " & Fn(Fx)
    X "Act ws : " & ActWsnn
    X NmvzDrs("Mis ws : ", IpxMisi)
    Stop
    'PushIAy O, AmTab(XX)
Next
VfxWsMis = O
'Insp "QDao_Lnk_LnkEr.LnkEr", "Inspect", "Oup(VfxWsMis) ExWsMis IpxMis ActWs",ExWsMis, ExWsMis, FmtDrs(IpxMis), FmtDrs(ActWs): Stop
End Function

Private Sub LnkErzS__Tst()
Brw LnkErzS(SampLnkImpSrc)
End Sub
Private Function VinpfDupEr() As InpfDupEu()

End Function

Private Function VinpnDupEu() As InpnDupEu()
'@Ipf : L FilKd Ffn IsFx HasFfn @@
'Dim Ffn$(): Ffn = StrCol(Ipf, "Ffn")
'Dim Dup$(): Dup = AwDup(Ffn)
'If Si(Dup) = 0 Then Exit Function
'Dim DupD As Drs: DupD = DwIn(Ipf, "Ffn", Dup)
'XBox "Ffn Duplicated"
'XDrs DupD
'XLin
'Stop
End Function

Private Function VinpfMisEr() As InpfMisEu()
'@Ipf : L FilKd Ffn IsFx HasFfn @@
'If NoReczDrs() Then Exit Function
'Dim A As Drs: A = DwEq(Ipf, "HasFfn", True) '! L Inp Ffn IsFx HasFfn
'Dim B As Drs: B = Vinp_AddCol_Pth_Fn(A)
'Dim C As Drs: C = SelDrs(B, "L FilKd Pth Fn")
'      VinpFfnMis = NmvzDrsO("File missing: ", C, DrsFmto(MaxWdt:=200))

'Insp "QDao_Lnk_LnkEr.LnkEr", "Inspect", "Oup(VinpFfnMis) EiFfnMis Ipf",EiFfnMis, EiFfnMis, FmtDrs(Ipf): Stop
End Function

Private Function VinpfKdDupEu() As InpfKdDupEu()
'@Ipf : L FilKd Ffn IsFx HasFfn @@
'Dim FilKd$(): FilKd = StrCol(Ipf, "FilKd")
'Dim Dup$(): Dup = AwDup(FilKd)
'If Si(Dup) = 0 Then Exit Function
'Dim DupD As Drs: DupD = DwIn(Ipf, "FilKd", Dup)
'XBox "FilKd Duplicated"
'XDrs DupD
'XLin
'VinpFilKdDup = XX
'Insp "QDao_Lnk_LnkEr.LnkEr", "Inspect", "Oup(VinpFilKdDup) EiFilKdDup Ipf",EiFilKdDup, EiFilKdDup, FmtDrs(Ipf): Stop
End Function

Private Function VfxTblDup() As FxTblDupEu()
End Function

Private Function VfxnDupEu() As FxnDupEu()
End Function
Private Function VfxnMis() As FxnMisEu()
End Function
Private Function VfxStruMisEu() As FxStruMisEu()
End Function
Private Function VfbnDupEu() As FbnDupEu()
End Function
Private Function VfbFbnMisUd() As FbnMisEu()

End Function
Private Function VfbnMis() As FbnMisEu()
End Function

Private Function VfbTblDup() As String()
End Function

Private Function VfbTblMis() As String()
End Function

Private Function VstruFldDupEu() As StruFldDupEu()
End Function

Private Function VstruTyErEu() As StruTyErEu()
End Function

Private Function VinpInptMis() As String()

End Function

Private Function VinpfMisEu(S As Lnkis) As InpEu
End Function


'---=============================================
Private Function VstruDup(IpsHdStru$()) As String()
'@IpsHdStru :  ! the stru coming from the Ips hd @@
'Insp "QDao_Lnk_LnkEr.LnkEr", "Inspect", "Oup(VstruDup) EsSDup IpsHdStru",EsSDup, EsSDup, IpsHdStru: Stop
End Function
Private Function VstruMisEu() As StruMisEu()
End Function
Private Function VstruExaEu() As StruExaEu()
End Function
Private Function VstruNoFldEr() As String()
End Function
Private Function VbexpTblExaEr() As String()
End Function
Private Function VbexpTblDupEr() As String()
End Function
Private Function VbexpTblMis(Ipw As Drs, Tny$()) As String()
'Fm:Wh@Ipw::Drs{L T Bexp}
Dim OL&(), OT$(), J%, T, Dr, O$()
For Each Dr In Itr(Ipw.Dy)
    T = Dr(1)
    If Not HasEle(Tny, T) Then
        PushI OL, Dr(0)
        PushI OT, T
    End If
Next
'---=
If Si(OL) = 0 Then Exit Function
For J = 0 To UB(OL)
    PushI O, FmtQQ("L#(?) Tbl(?) is not defined.", OL(J), OT(J))
Next
PushI O, vbTab & "Defined tables are:"

For Each T In Itr(Tny)
    PushI O, vbTab & vbTab & T
Next
VbexpTblMis = O
End Function
Private Function VbexpBexpEmpEr(Ipw As Drs) As String()
'Ret            : with tbl nm but no Bexp @@
Dim J%, OL&(), OT$(), O$()
'Fm : Wh@Ipw::Drs{L T Bexp}
Dim Dr, L&, T$, Bexp$
For Each Dr In Itr(Ipw.Dy)
    Bexp = Dr(2)
    If Bexp = "" Then
        L = Dr(0)
        T = Dr(1)
        PushI OL, L
        PushI OT, T
    End If
Next
'---
For J = 0 To UB(OL)
    PushI O, FmtQQ("L#(?) Tbl(?) has no Bexp", OL(J), OT(J))
Next
VbexpBexpEmpEr = O
'Insp "QDao_Lnk_LnkEr.LnkEr", "Inspect", "Oup(VbexpBexpEmp) EwBexpEmp Ipw",EwBexpEmp, EwBexpEmp, FmtDrs(Ipw): Stop
End Function

