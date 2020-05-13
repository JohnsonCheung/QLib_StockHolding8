Attribute VB_Name = "MxIdeSrcLis"
Option Compare Text
Option Explicit
Const CNs$ = "Src.Lis"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcLis."

Private Sub BrwSrc__Tst()
BrwSrc "Dim J%"
End Sub

Sub DmpSrc(Patn$): LisSrc Patn, Dmpg: End Sub
Sub VcSrc(Patn$): LisSrc Patn, Vcg: End Sub
Sub BrwSrc(Patn$): LisSrc Patn, Brwg("Jrc_"): End Sub
Private Sub LisSrc(Patn$, Oup As OupOpt): LisAy Jrc(CPj, Patn), Oup: End Sub
