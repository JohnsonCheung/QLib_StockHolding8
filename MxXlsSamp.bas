Attribute VB_Name = "MxXlsSamp"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsSamp."
Function SampLo() As ListObject
Set SampLo = NwLo(RgzSq(SampSqWithHdr, NwA1), "Sample")
End Function

Function SampLoVis() As ListObject
Set SampLoVis = VisLo(SampLo)
End Function

Function SampPt() As PivotTable
Set SampPt = PtzRg(SampRg)
End Function
Function SampRg() As Range
Set SampRg = VisRg(NwLozSq(SampSq, NwA1))
End Function

Function SampLofTp() As String()
Dim O$()
PushI O, "Lo  Nm     *Nm"
PushI O, "Lo  Fld    *Fld.."
PushI O, "Ali Left   *Fld.."
PushI O, "Ali Right  *Fld.."
PushI O, "Ali Center *Fld.."
PushI O, "Bdr Left   *Fld.."
PushI O, "Bdr Right  *Fld.."
PushI O, "Bdr Col    *Fld.."
PushI O, "Tot Sum    *Fld.."
PushI O, "Tot Avg    *Fld.."
PushI O, "Tot Cnt    *Fld.."
PushI O, "Fmt *Fmt   *Fld.."
PushI O, "Wdt *Wdt   *Fld.."
PushI O, "Lvl *Lvl   *Fld.."
PushI O, "Cor *Cor   *Fld.."
PushI O, "Fml *Fld   *Formula"
PushI O, "Bet *Fld   *Fld1 *Fld2"
PushI O, "Tit *Fld   *Tit"
PushI O, "Lbl *Fld   *Lbl"
SampLofTp = O
End Function

Function SampWs() As Worksheet
Dim O As Worksheet
Set O = NwWs
LozDrs SampDrs, WsRC(O, 2, 2)
Set SampWs = O
VisWs O
End Function
