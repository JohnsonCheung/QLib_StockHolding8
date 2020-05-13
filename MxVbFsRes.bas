Attribute VB_Name = "MxVbFsRes"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVbFsRes."

Function EnsResPseg(Pseg$)
EnsAllFdr ResPth(Pseg)
End Function

Function ResHom$()
Static H$: If H = "" Then H = AddFdrEns(AssPthP, ".res")
ResHom = H
End Function

Function ResPth$(Pseg$)
ResPth = EnsPthSfx(ResHom & Pseg)
End Function

Function ResFfn$(SegFn$)
ResFfn = ResHom & SegFn
End Function

Function ResFfnEns$(SegFn$)
Dim O$: O = ResFfn(SegFn)
ResFfnEns = EnsAllFdr(Pth(O))
End Function

Sub WrtRes(S$, SegFn$, Optional OvrWrt As Boolean)
WrtStr S, ResFfn(SegFn), OvrWrt
End Sub

Function Resl$(SegFn$)
Resl = LineszFt(ResFfn(SegFn))
End Function

Function ResLy(SegFn$) As String()
ResLy = SplitCrLf(Resl(SegFn))
End Function

Function ResDrs(SegFn$) As Drs
ResDrs = DrszFt(ResFfn(SegFn))
End Function

Function ResLo(SegFn$) As ListObject
Dim F$: F = ResFfn(SegFn)
OpnFcsv F
Set ResLo = NwLo(DtaRg(FstWs(LasCWb)))
End Function

Sub ResDrs__Tst()
Dim D As Drs: D = ResDrs("MthDrsP")
Stop
End Sub

Sub WrtResDrs(D As Drs, SegFn$)
WrtDrs D, ResFfn(SegFn)
End Sub

Function ResS12y(SegFn$) As S12()
ResS12y = S12y(ResLy(SegFn))
End Function
