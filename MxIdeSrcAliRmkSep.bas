Attribute VB_Name = "MxIdeSrcAliRmkSep"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Ali.RmkSep"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcAliRmkSep."
Function SelRmkSep(WiMthln As Drs) As Drs
'@WiMthln : Mthln #Mth-Context.
'Ret : select where LTrim-*Mthln has pfx '-- '== or '..
Dim IxMthln%: IxMthln = IxzAy(WiMthln.Fny, "L")
Dim Dr, Dy(): For Each Dr In Itr(WiMthln.Dy)
    Dim L$: L = LTrim(Dr(IxMthln))
    If FstChr(L) = "'" Then
        L = Left(RmvFstChr(L), 2)
        Select Case L
        Case "==", "--", "..": PushI Dy, Dr
        End Select
    End If
Next
SelRmkSep.Fny = WiMthln.Fny
SelRmkSep.Dy = Dy
End Function

Function AliRmkSepzD(Wi_L_Mthln As Drs) As Drs
'@Wi_L_Mthln De : L Mthln ! Where Mthln is {spc}'-- '== '..
'Ret   : L NewL OldL        ! Where NewL is aligned with 120 @@
Dim IxL%, IxMthln%: AsgIx Wi_L_Mthln, "L Mthln", IxL, IxMthln
Dim Dr, Dy(): For Each Dr In Itr(Wi_L_Mthln.Dy)
    Dim L&:       L = Dr(IxL)
    Dim Oldl$: Oldl = Dr(IxMthln)
    Dim C$:       C = Mid(LTrim(Oldl), 2, 1)
    Dim Newl$: Newl = Left(Oldl, 120) & Dup(C, 120 - Len(Oldl))
    If Oldl <> Newl Then
        Push Dy, Array(L, Newl, Oldl)
    End If
Next
AliRmkSepzD = LNewO(Dy)
'Insp "QIde_B_AliMth.XDeLNewO", "Inspect", "Oup(XDeLNewO) De", FmtDrs(XDeLNewO), FmtDrs(De): Stop
End Function

Function AliRmkSep(Upd As Boolean, M As CodeModule, Wi_L_Mthln As Drs) As Drs
Dim D As Drs:   D = SelRmkSep(Wi_L_Mthln)
Dim D1 As Drs: D1 = AliRmkSepzD(D)
:                    If Upd Then RplLNewO M, D ' <==
End Function
