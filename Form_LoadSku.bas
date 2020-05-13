VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LoadSku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = CLib & "Form_LoadSku."

Private Sub BusAreaOvr_BeforeUpdate(Cancel As Integer)
'DoCmd.RunSQL "Update Sku set BusArea=IIf(Trim(Nz(BusAreaOvr,''))='',BusAreaSap,BusAreaOvr) where Sku='" & Me.Sku & "'"
Me.BusArea.Value = Me.BusAreaOvr.Value
End Sub

Private Sub Litre_per_Btl_Ovr_AfterUpdate()
RunCQ "Update Sku set [Litre/Btl]=IIf(Trim(Nz([Litre/Btl],''))='',[Litre/BtlSap],[Litre/BtlOvr]) where Sku='" & Me.SKU & "'"
End Sub

Private Sub Cmd_Sel_IFx_Click():  SelCFxPm "SkuIFx", "Select Sales Text.Xlsx file", Me.xFx: End Sub
Private Sub Cmd_Opn_IPth_Click(): BrwPth Pth(SkuPrm_InpFx): End Sub

Private Sub Cmd_Rpt_SkuL_Click():     RptSkuL:               End Sub
Private Sub Cmd_Load_Sku_Click():     LoadSku:               End Sub
Private Sub Cmd_Opn_IFx_Click():      OpnFx SkuPrm_InpFx:   End Sub
Private Sub Cmd_Sel_CpyToPth_Click(): FrmSkuL_Sel_CpyToPth:  End Sub
Private Sub Cmd_Opn_OPth_Click():     BrwPth AppOPth:    End Sub
Private Sub Cmd_Tgl_IsCpyTo_Click():  FrmSkuL_Tgl_IsCpyTo: End Sub

Private Sub Cmd_Exit_Click():              DoCmd.Close:           End Sub
Private Sub Form_Close():                  RfhTbSku_Ovr: End Sub
Private Sub Form_Open(Cancel As Integer):  FrmSkuL_FrmOpened: End Sub
