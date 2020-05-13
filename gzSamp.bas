Attribute VB_Name = "gzSamp"
Option Explicit
Option Compare Text
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzSamp."
Function SampWhereDec24$()
SampWhereDec24 = OHYmdBexp(SampDDec24)
End Function
Function SampDNov28() As Ymd
SampDNov28 = Ymd(19, 11, 28)
End Function
Function SampDDec24() As Ymd
SampDDec24 = Ymd(19, 12, 24)
End Function
Function Samp86Dec24() As CoYmd
Samp86Dec24 = CoYmd(86, SampDDec24)
End Function
