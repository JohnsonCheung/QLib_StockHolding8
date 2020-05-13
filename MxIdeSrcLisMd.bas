Attribute VB_Name = "MxIdeSrcLisMd"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcLisMd."

Private Sub BrwMd__Tst()
BrwMd
End Sub

Sub LisQVb()
LisMd Lib:="QVb", SrtFF:="-NMth", NsPatn:="Str"
End Sub

Sub BrwMd(Optional MdPatn$, Optional NsPatn$, Optional Lib$, Optional SrtFF$)
LisMd MdPatn, NsPatn, Lib, SrtFF, OupTy:=eBrwOup, Top:=0
End Sub

Sub DmpMd(Optional MdPatn$, Optional NsPatn$, Optional Lib$, Optional SrtFF$, Optional Top% = 50)
LisMd MdPatn, NsPatn, Lib, SrtFF, Top, eDmpOup
End Sub
Sub LisMd(Optional MdPatn$, Optional NsPatn$, Optional Lib$, Optional SrtFF$, Optional Top% = 50, Optional OupTy As eOupTy = eOupTy.eDmpOup)
Dim D1 As Drs: D1 = MdDrsP(MdPatn)
Dim D2 As Drs: D2 = DwPatn(D1, "CNsv", NsPatn)
                    If Lib <> "" Then D2 = DwEq(D2, "CLibv", Lib)
Dim D3 As Drs: D3 = SrtDrs(D2, SrtFF)
'Insp "LisMd", "The 4 Drs", "MdPatn D1 D2 D3", CsvLy(MdDrsP), CsvLy(D1), CsvLy(D2), CsvLy(D3): Stop
Dim Ly$():     Ly = FmtDrsR(D3)
                    LisAy Ly, OupOpt("LisMd_", OupTy)
End Sub

Sub VcMd(Optional MdPatn$, Optional NsPatn$, Optional Lib$, Optional SrtFF$)
LisMd MdPatn, NsPatn, Lib, SrtFF, OupTy:=eVcOup, Top:=0
End Sub
