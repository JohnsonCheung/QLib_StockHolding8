Attribute VB_Name = "MxAcsSrc"
Option Compare Text
Option Explicit
Const CNs$ = "Acs.Src"
Const CLib$ = "QAcs."
Const CMod$ = CLib & "MxAcsSrc."
Function SrclzAcs$(A As Access.Application, Mdn)
SrclzAcs = Srcl(PjzAcs(A).VBComponents(Mdn))
End Function
