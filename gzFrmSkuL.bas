Attribute VB_Name = "gzFrmSkuL"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzFrmSkuL."
Sub RptSkuL()
Dim A$: A = Format(LasOHDte, "YYYY-MM-DD")
Dim Fx$: Fx = CDbPth & "Output\OnHand SkuLis(" & A & ") Gen@" & Format(Now, "YYYY-MM-DD HHMM") & ".xlsx"
Dim Tp$: Tp = CDbPth & "WorkingDir\Templates\Sku List Template.xlsx"
RfhTbSku_WithOHxxx
RunCQ "Select * into [@Sku] from [qSku]"
Dim Wb As Workbook: Set Wb = RfhFx(Fx, Tp, CFb)
Dim CpyToPth$: CpyToPth = SkuLisPrm_CpyToPth
If SkuLisPrm_IsCpyTo Then
    Dim Fso As New Scripting.FileSystemObject
    Fso.CopyFile Fx, CpyToPth, True
End If
End Sub
'---
Private Sub FrmSkuL_Tgl_IsCpyTo__Tst()
Debug.Print SkuLisPrm_IsCpyTo
FrmSkuL_Tgl_IsCpyTo
Debug.Print SkuLisPrm_IsCpyTo
End Sub

Sub FrmSkuL_Tgl_IsCpyTo()
SetCPmv "SkuLis_IsCpyTo", Not SkuLisPrm_IsCpyTo:
Set_IsCpyToVis
End Sub

Sub FrmSkuL_FrmOpened()
DoCmd.Maximize
RfhTbSku_WithOHxxx
Set_IsCpyToVis
RfhTbSku_Ovr
End Sub
Sub FrmSkuL_Sel_CpyToPth()
SelCPthPm "SkuLis_CpyToPth":
Set_IsCpyToVis
End Sub
Private Sub Set_IsCpyToVis__Tst()
Set_IsCpyToVis
End Sub
Private Sub Set_IsCpyToVis()
Dim V As Boolean: V = SkuLisPrm_IsCpyTo
F.Cmd_Sel_CpyToPth.Visible = V
F.xCpyToPth.Visible = V
F.Cmd_Opn_CpyToPth.Visible = V
End Sub

Private Function F() As Form_LoadSku: Set F = Form_LoadSku: End Function
