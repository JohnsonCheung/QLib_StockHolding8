Attribute VB_Name = "MxIdeSrcUtfSig"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxIdeSrcUtfSig."
Public Const Utf8Sig$ = "﻿"

Function RmvUtf8Sig$(S$)
RmvUtf8Sig = RmvPfx(S, Utf8Sig)
End Function

Private Sub HasUtfSig8__Tst()
Dim F$: F = LineszFt(ResFfn("MthDrsP\"))
Debug.Assert HasUtf8Sig(F)
End Sub

Function HasUtf8Sig(Ft$) As Boolean
HasUtf8Sig = HasPfx(FstNChrzFfn(Ft, 3), Utf8Sig, vbBinaryCompare)
End Function
