Attribute VB_Name = "MxIdeSrcLisMdn"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Lis"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcLisMdn."

Sub BrwMdn(Optional MdPatn$, Optional NsPatn$, Optional Lib$, Optional SrtFF$)
LisMdn MdPatn, NsPatn, Lib, SrtFF, eBrwOup, Top:=0
End Sub

Sub LisMdn(Optional MdPatn$, Optional NsPatn$, Optional Lib$, Optional SrtFF$, Optional T As eOupTy, Optional Top% = 50)
Dim D1 As Drs: D1 = DwPatn(MdnDrsP, "Mdn", MdPatn)
Dim D2 As Drs: D2 = DwPatn(D1, "CNsv", NsPatn)
                    If Lib <> "" Then D2 = DwEq(D2, "CLibv", Lib)
Dim D3 As Drs: D3 = SrtDrs(D2, SrtFF)
Dim Ly$():     Ly = FmtDrsR(D3)
                    LisAy Ly, OupOpt("LisMdn_", T)
End Sub

Sub VcMdn(Optional MdPatn$, Optional NsPatn$, Optional Lib$, Optional SrtFF$)
LisMdn MdPatn, NsPatn, Lib, SrtFF, eVcOup, Top:=0
End Sub
