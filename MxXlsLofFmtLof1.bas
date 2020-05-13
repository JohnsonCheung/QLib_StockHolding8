Attribute VB_Name = "MxXlsLofFmtLof1"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsLofFmtLof1."
Const CNs$ = "Lof"
Const LofLonC$ = "Lon ?"
Const LofFnyC$ = "Fny ?"
Const LofAliC$ = "Bdr ? ?"
Const LofBdrC$ = "Bdr ? ?"
Const LofCorC$ = "Cor ? ?"
Const LofFmlC$ = "Fml ? ?"
Const LofFmtC$ = "Fmt ? ?"
Const LofLblC$ = "Lbl ? ?"
Const LofLvlC$ = "Lvl ? ?"
Const LofSumC$ = "Sum ? ?"
Const LofTitC$ = "Tit ? ?"
Const LofTotC$ = "Tot ? ?"
Const LofWdtC$ = "Wdt ? ?"

Sub BrwSampLof(): Brw FmtLof(SampLof): End Sub

Function FmtLof(A As Lof) As String()
PushI FmtLof, FmtLofLon(A.Lon)
PushI FmtLof, FmtLofFny(A.Fny)
PushIAy FmtLof, FmtLofAli(A.Ali)
PushIAy FmtLof, FmtLofBdr(A.Bdr)
PushIAy FmtLof, FmtLofCor(A.Cor)
PushIAy FmtLof, FmtLofFml(A.Fml)
PushIAy FmtLof, FmtLofFmt(A.Fmt)
PushIAy FmtLof, FmtLofLbl(A.Lbl)
PushIAy FmtLof, FmtLofLvl(A.Lvl)
PushIAy FmtLof, FmtLofSum(A.Sum)
PushIAy FmtLof, FmtLofTit(A.Tit)
PushIAy FmtLof, FmtLofTot(A.Tot)
PushIAy FmtLof, FmtLofWdt(A.Wdt)
End Function

Private Function FmtLofLon$(Lon$): FmtLofLon = FmtQQ(LofLonC, Lon): End Function
Private Function FmtLofFny$(Fny$()): FmtLofFny = FmtQQ(LofFnyC, JnSpc(Fny)): End Function

Private Function FmtLofAli(A() As LofAli) As String(): Dim J%: For J = 0 To LofAliUB(A): PushI FmtLofAli, FmtItmAli(A(J)): Next: End Function
Private Function FmtLofBdr(A() As LofBdr) As String(): Dim J%: For J = 0 To LofBdrUB(A): PushI FmtLofBdr, FmtItmBdr(A(J)): Next: End Function
Private Function FmtLofCor(A() As LofCor) As String(): Dim J%: For J = 0 To LofCorUB(A): PushI FmtLofCor, FmtItmCor(A(J)): Next: End Function
Private Function FmtLofFml(A() As LofFml) As String(): Dim J%: For J = 0 To LofFmlUB(A): PushI FmtLofFml, FmtItmFml(A(J)): Next: End Function
Private Function FmtLofFmt(A() As LofFmt) As String(): Dim J%: For J = 0 To LofFmtUB(A): PushI FmtLofFmt, FmtItmFmt(A(J)): Next: End Function
Private Function FmtLofLbl(A() As LofLbl) As String(): Dim J%: For J = 0 To LofLblUB(A): PushI FmtLofLbl, FmtItmLbl(A(J)): Next: End Function
Private Function FmtLofLvl(A() As LofLvl) As String(): Dim J%: For J = 0 To LofLvlUB(A): PushI FmtLofLvl, FmtItmLvl(A(J)): Next: End Function
Private Function FmtLofSum(A() As LofSum) As String(): Dim J%: For J = 0 To LofSumUB(A): PushI FmtLofSum, FmtItmSum(A(J)): Next: End Function
Private Function FmtLofTit(A() As LofTit) As String(): Dim J%: For J = 0 To LofTitUB(A): PushI FmtLofTit, FmtItmTit(A(J)): Next: End Function
Private Function FmtLofTot(A() As LofTot) As String(): Dim J%: For J = 0 To LofTotUB(A): PushI FmtLofTot, FmtItmTot(A(J)): Next: End Function
Private Function FmtLofWdt(A() As LofWdt) As String(): Dim J%: For J = 0 To LofWdtUB(A): PushI FmtLofWdt, FmtItmWdt(A(J)): Next: End Function


Function LofAliUB&(A() As LofAli): LofAliUB = LofAliSi(A) - 1: End Function
Function LofBdrUB&(A() As LofBdr): LofBdrUB = LofBdrSi(A) - 1: End Function
Function LofCorUB&(A() As LofCor): LofCorUB = LofCorSi(A) - 1: End Function
Function LofFmlUB&(A() As LofFml): LofFmlUB = LofFmlSi(A) - 1: End Function
Function LofFmtUB&(A() As LofFmt): LofFmtUB = LofFmtSi(A) - 1: End Function
Function LofLblUB&(A() As LofLbl): LofLblUB = LofLblSi(A) - 1: End Function
Function LofLvlUB&(A() As LofLvl): LofLvlUB = LofLvlSi(A) - 1: End Function
Function LofSumUB&(A() As LofSum): LofSumUB = LofSumSi(A) - 1: End Function
Function LofTitUB&(A() As LofTit): LofTitUB = LofTitSi(A) - 1: End Function
Function LofTotUB&(A() As LofTot): LofTotUB = LofTotSi(A) - 1: End Function
Function LofWdtUB&(A() As LofWdt): LofWdtUB = LofWdtSi(A) - 1: End Function

Function LofAliSi&(A() As LofAli): On Error Resume Next: LofAliSi = UBound(A) + 1: End Function
Function LofBdrSi&(A() As LofBdr): On Error Resume Next: LofBdrSi = UBound(A) + 1: End Function
Function LofCorSi&(A() As LofCor): On Error Resume Next: LofCorSi = UBound(A) + 1: End Function
Function LofFmlSi&(A() As LofFml): On Error Resume Next: LofFmlSi = UBound(A) + 1: End Function
Function LofFmtSi&(A() As LofFmt): On Error Resume Next: LofFmtSi = UBound(A) + 1: End Function
Function LofLblSi&(A() As LofLbl): On Error Resume Next: LofLblSi = UBound(A) + 1: End Function
Function LofLvlSi&(A() As LofLvl): On Error Resume Next: LofLvlSi = UBound(A) + 1: End Function
Function LofSumSi&(A() As LofSum): On Error Resume Next: LofSumSi = UBound(A) + 1: End Function
Function LofTitSi&(A() As LofTit): On Error Resume Next: LofTitSi = UBound(A) + 1: End Function
Function LofTotSi&(A() As LofTot): On Error Resume Next: LofTotSi = UBound(A) + 1: End Function
Function LofWdtSi&(A() As LofWdt): On Error Resume Next: LofWdtSi = UBound(A) + 1: End Function

Function FmtItmAli$(A As LofAli): FmtItmAli = FmtQQ(LofAliC, LofAliStr(A.Ali), JnSpc(A.Fny)): End Function
Function FmtItmBdr$(A As LofBdr): FmtItmBdr = FmtQQ(LofBdrC, LofBdrStr(A.Bdr), JnSpc(A.Fny)): End Function
Function FmtItmCor$(A As LofCor): FmtItmCor = FmtQQ(LofCorC, LofCorStr(A.Cor), JnSpc(A.Fny)): End Function
Function FmtItmFml$(A As LofFml): FmtItmFml = FmtQQ(LofFmlC, A.Fld, A.Fml):                   End Function
Function FmtItmFmt$(A As LofFmt): FmtItmFmt = FmtQQ(LofFmtC, A.Fmt, JnSpc(A.Fny)):            End Function
Function FmtItmLbl$(A As LofLbl): FmtItmLbl = FmtQQ(LofLblC, A.Fld, A.Lbl):                   End Function
Function FmtItmLvl$(A As LofLvl): FmtItmLvl = FmtQQ(LofLvlC, A.Lvl, JnSpc(A.Fny)):            End Function
Function FmtItmSum$(A As LofSum): FmtItmSum = FmtQQ(LofSumC, A.SumFld, A.FmFld, A.ToFld):     End Function
Function FmtItmTit$(A As LofTit): FmtItmTit = FmtQQ(LofAliC, A.Fld, A.Tit):                   End Function
Function FmtItmTot$(A As LofTot): FmtItmTot = FmtQQ(LofAliC, LofTotStr(A.Tot), JnSpc(A.Fny)): End Function
Function FmtItmWdt$(A As LofWdt): FmtItmWdt = FmtQQ(LofAliC, A.Wdt, JnSpc(A.Fny)):            End Function

Function LofAliStr$(A As eLofAli)

End Function

Function LofBdrStr$(A As eLofBdr)

End Function

Function LofCorStr$(Colr&)

End Function


Function LofTotStr$(A As eLofTot)

End Function
