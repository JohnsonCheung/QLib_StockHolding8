Attribute VB_Name = "MxVbRunTmpCd"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVbRunTmpCd."

Sub RunCdLy(CdLy$())
RunCd JnCrLf(CdLy)
End Sub

Sub RunCd(CdLines$)
Dim N$: N = TmpNm("TmpMth_")
AddMthByCd TmpCdMd, N, CdLines
Run N
End Sub

Function TmpCdMd() As CodeModule
Const Mdn$ = "ZTmpModForTmpCd"
EnsMod CPj, Mdn
Set TmpCdMd = Md(Mdn)
End Function
