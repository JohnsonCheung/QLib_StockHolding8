Attribute VB_Name = "MxIdeSrcLisJrc"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcLisJrc."
Public Const MthFFLis$ = "Pjn MdTy Mdn L Mdy Ty Mthn TyChr RetAs ShtPm"
'JSrc:: :Ln #Jmp-Ln# ! Fmt: T1 Rst, *T1 is JmpLin"<mdn:Lno>".  *Rst is '<SrcLin>

Sub LisPj()
Dim A$()
    A = PjnyzV(CVbe)
    D AmAddPfx(A, "ShwPj """)
D A
End Sub

Sub LisStopLn()
LisJrc "Stop"
End Sub

Sub LisJrc(LnPatn$, Optional Oup As eOupTy = eOupTy.eDmpOup) ' List jump source line of CPj using @LnPatn and @Oup
'Jrc:Cml #Jmp-Src-Ln# A source line which can jump to that SrcLn and that SrcLn will the remark
LisJrczP CPj, LnPatn$, Oup
End Sub

Sub LisJrczP(P As VBProject, LnPatn$, Optional OupTy As eOupTy = eOupTy.eDmpOup)
LisAy Jrc(P, LnPatn), OupOpt("LisJrc_", OupTy)
End Sub

Sub LisJrczPfx(LinPfx$, Optional T As eOupTy = eOupTy.eDmpOup)
LisAy JrcyzPfx(LinPfx), OupOpt("LisJrczPfx_", T)
End Sub

Sub LisJrczIdr(Idr$, Optional Oup As eOupTy = eOupTy.eDmpOup) ' List jump source lines using Identitifier and @Oupt
'Idr:Cml :Nm #Identifier#  A name at begin of a line or begin of spc or opnBkt
LisAy JrcyzIdr(Idr), OupOpt("LisJrczIdr_", Oup)
End Sub

Function PatnzWhSS$(WhSS, LisAy$()) ' ret a Patn of XX|XX..|XX where XX is from @WhSS after normalized by @LisAy
Dim WhSy$(): WhSy = AwDis(SyzSS(WhSS))
Dim Wh$(): Wh = IntersectAy(WhSy, LisAy)
PatnzWhSS = Jn(Wh, "|")
End Function

Private Sub Beta__Tst()
'ß()
End Sub
