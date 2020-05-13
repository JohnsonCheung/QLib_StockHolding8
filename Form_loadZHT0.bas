VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_loadZHT0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = CLib & "Form_loadZHT0."
Option Base 0
Private Const SelSql$ = "SELECT SKU, SkuDes, TaxRateHK, TaxRateMO, TaxUOMHK, TaxUOMMO," & _
" DteCrt , DteUpdTaxRate, WithOHCur, WithOHHst FROM SKU"
Private Const OrdBy$ = " Order by Sku"
Private Const WhAll$ = " Where Nz(TaxRateHK,0)<>0 or Nz(TaxRateMO,0)<>0"
Private Const WhHK$ = " Where Nz(TaxRateHK,0)<>0"
Private Const WhMO$ = " Where Nz(TaxRateMO,0)<>0"
Private Const ShwHKSql$ = SelSql & WhHK & OrdBy
Private Const ShwMOSql$ = SelSql & WhMO & OrdBy
Private Const ShwAllSql$ = SelSql & WhAll & OrdBy
Private Sub Cmd_Exit_Click():       DoCmd.Close: End Sub

Private Sub Cmd_UpLoad_Click():  LoadZHT0: Requery:  End Sub
Private Sub Cmd_Opn_IFx_Click(): OpnFx ZHT0IFx: End Sub
Private Sub Cmd_Opn_WFx_Click(): OpnFx ZHT0WFx: End Sub
Private Sub Cmd_Opn_IPth_Click(): BrwPth ZHT0IPth: End Sub
Private Sub Cmd_Sel_IFx_Click(): SelCFxPm "ZHT0_InpFx", , xInpFx: End Sub

Private Sub Cmd_Shw_All_Click(): ShwAll: End Sub
Private Sub Cmd_Shw_HK_Click():  ShwHK:  End Sub
Private Sub Cmd_Shw_MO_Click():  ShwMO:  End Sub
Private Sub ShwAll(): Me.RecordSource = ShwAllSql: Requery: End Sub
Private Sub ShwHK(): Me.RecordSource = ShwHKSql:   Requery: End Sub
Private Sub ShwMO(): Me.RecordSource = ShwMOSql:   Requery: End Sub

Private Sub Form_Load()
Me.RecordSource = ShwAllSql
End Sub
