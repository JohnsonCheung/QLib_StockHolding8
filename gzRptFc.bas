Attribute VB_Name = "gzRptFc"
Option Explicit
Option Compare Text
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzRptFc."
Private Sub RptFc__Tst()
RptFc YM(19, 12)
End Sub
Sub RptFc(A As YM)
Dim OFx$: OFx = FcOFx(A): If OpnFxIfExist(OFx) Then Exit Sub
DltFfn OFx
TmpFc_ByYM A
TmpPH5
AddPH7Atr "$Fc?", "@Fc?"
SrtFldPH7 "@Fc?", ExpandPfxNN("M", 1, 15, "00")
MaxiWb RfhFx(OFx, FcTp, CFb)
End Sub

Function FcOFx$(A As YM)
FcOFx = FcOPth & FcOFxFn(A)
End Function
Function FcOFxFn$(A As YM)
FcOFxFn = "Forecast " & YymStr(A.Y, A.M) & ".xlsx"
End Function
Function FcTp$()
FcTp = TpPthP & "Forecast (Template).xlsx"
End Function
