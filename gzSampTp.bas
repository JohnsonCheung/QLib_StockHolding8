Attribute VB_Name = "gzSampTp"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzSampTp."
Function SampShFx$(): SampShFx = ShOPth & "Sample Stock Holding Report.xlsx": End Function
Private Sub EnsSampShFx(): EnsFfnFm SampShFx, ShTp: End Sub
Function SampShWb() As Workbook
EnsSampShFx
Static Wb As Workbook
If Not IsGoodWb(Wb) Then Set Wb = WbzFx(SampShFx)
Set SampShWb = Wb
End Function

Function SampShLo() As ListObject: Set SampShLo = FstLo(SampShWs): End Function
Function SampSdWs() As Worksheet: Set SampSdWs = SampShWb.Sheets("StkDays Stm"): SampSdWs.Activate: End Function
Function SampFcWs() As Worksheet: Set SampFcWs = SampShWb.Sheets("Fc Stm"):      SampFcWs.Activate: End Function
Function SampShWs() As Worksheet: Set SampSdWs = SampShWb.Sheets("StkHld Stm"): SampSdWs.Activate: End Function
