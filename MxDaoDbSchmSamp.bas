Attribute VB_Name = "MxDaoDbSchmSamp"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDbSchmSamp."

Function SampSchmSrc(N%) As SchmSrc: SampSchmSrc = SchmSrczS(SampSchm(N)): End Function

Function SampSchmFt$(N%)
SampSchmFt = ResFfn(ZZFn(N))
End Function

Function SampSchm(N%) As String()
SampSchm = ResLy(ZZFn(N))
End Function

Sub EdtSampSchm(N%)
EdtRes ZZFn(N)
End Sub

Private Function ZZFn$(N%)
ZZFn = "SampSchm" & N & ".schm.txt"
End Function
