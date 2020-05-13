Attribute VB_Name = "MxDtaDaBrw"
Option Explicit
Option Compare Text
Const CNs$ = "Vb.Brw"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaBrw."

Sub BrwDrs2(A As Drs, B As Drs, Optional NN$ = "Drs1 Drs2", Optional Tit$ = "Brw 2 Drs")
Dim T$(), AyA$(), AyB$(), N1$, N2$
N1 = BefSpc(NN)
N2 = AftSpc(NN)
AyA = FmtDrsR(A)
AyB = FmtDrsR(B)
T = UL(Tit)
BrwAy Sy(T, AyA, AyB), "BrwDrs2_"
End Sub

Sub BrwDrs3(A As Drs, B As Drs, C As Drs, Optional ByVal NN$ = "Drs1 Drs2 Drs3", Optional Tit$ = "Brw 3 Drs")
Dim T$(), AyA$(), AyB$(), AyC$(), N1$, N2$, N3$
N1 = ShfTerm(NN)
N2 = ShfTerm(NN)
N3 = NN
AyA = FmtDrsR(A, N1)
AyB = FmtDrsR(B, N2)
AyC = FmtDrsR(C, N3)
T = UL(Tit)
BrwAy Sy(T, AyA, AyB, AyC), "BrwDrs3_"
End Sub

Sub BrwDrs4(A As Drs, B As Drs, C As Drs, D As Drs, Optional ByVal NN$ = "Drs1 Drs2 Drs3 Drs4", Optional Tit$ = "Brw 4 Drs")
Dim T$(), AyA$(), AyB$(), AyC$(), AyD$(), N1$, N2$, N3$, N4$
N1 = ShfTerm(NN)
N2 = ShfTerm(NN)
N3 = ShfTerm(NN)
N4 = NN
AyA = FmtDrsR(A, N1)
AyB = FmtDrsR(B, N2)
AyC = FmtDrsR(C, N3)
AyD = FmtDrsR(D, N4)
T = UL(Tit)
BrwAy Sy(T, AyA, AyB, AyC, AyD), "BrwDrs4_"
End Sub

Sub BrwDrs(D As Drs, Optional FnPfx$) ' browse Drs reducely with default format option
BrwDrsR D, FnPfx
End Sub

Sub BrwDrsR(D As Drs, Optional FnPfx$) ' browse Drs reducely with default format option
':BrwDrs: :Fun #Brw-Drs-Normally#
LisAy FmtDrsRO(D, DftDrsFmto), Brwg(FnPfx)
End Sub

Sub BrwDrsNO(D As Drs, Opt As DrsFmto, Optional FnPfx$) ' browse Drs normal (without reduce) with option
':BrwDrsN: :Fun #Brw-Drs-Normally#
LisAy FmtDrsNO(D, Opt), Brwg(FnPfx)
End Sub

Sub BrwDrsN(D As Drs, Optional FnPfx$ = "Drs_") ' browse Drs normal (without reduce)
BrwDrsNO D, DftDrsFmto, FnPfx
End Sub

Sub BrwDrsRO(D As Drs, Opt As DrsFmto, Optional FnPfx$ = "Drs_") ' browse Drs @D reducely with option
LisAy FmtDrsRO(D, Opt), Brwg(FnPfx)
End Sub

Sub VcDrsO(D As Drs, Opt As DrsFmto, Optional FnPfx$ = "Drs_")
VcAy FmtDrsRO(D, Opt), FnPfx
End Sub

Sub VcDrs(D As Drs)
VcDrsO D, DftDrsFmto
End Sub

Sub BrwDt(D As Dt, Optional FnPfx$): BrwAy FmtDt(D), FnPfx: End Sub

Sub BrwDy(D(), Optional MaxColWdt% = 100, Optional Fmt As eTblFmt = eTblFmt.eNoSep)
BrwAy FmtDy(D, MaxColWdt, Fmt)
End Sub

Sub DmpDy(Dy(), _
Optional MaxColWdt% = 100, _
Optional Fmt As eTblFmt)
D FmtDy(Dy, MaxColWdt, Fmt)
End Sub
Sub DmpDrsO(D As Drs, Opt As DrsFmto) ' dump a @D reducely with format optional
DmpAy FmtDrsRO(D, Opt)
End Sub

Sub BrwDsRO(D As Ds, Opt As DrsFmto, Optional FnPfx$ = "BrwDs_") ' browse dataset @D using default format option
BrwAy FmtDsRO(D, Opt, FnPfx)
End Sub

Function FmtDsRO(D As Ds, Opt As DrsFmto, Optional FnPfx$ = "BrsDs_") As String()
Dim J%: For J = 0 To DtUB(D.DtAy)
    Dim Dt As Dt: Dt = D.DtAy(J)
    PushIAy FmtDsRO, FmtDrsRO(DrszDt(Dt), Opt, Dt.DtNm)
Next
End Function

Sub BrwDs(D As Ds, Optional MaxWdt% = 100, Optional Brkcc$, Optional ShwZer As Boolean, Optional IsSum As Boolean, Optional BegIx%, Optional Fmt As eTblFmt, Optional FnPfx$ = "BrwDrs_") ' browse dataset @D using default format option
BrwAy FmtDsO(D, DrsFmto(MaxWdt, Brkcc, ShwZer, IsSum, BegIx, Fmt)), FnPfx
End Sub

Sub DmpDs(D As Ds)
DmpAy FmtDs(D)
End Sub
