Attribute VB_Name = "gzRfhTbSku_WithOH"
Option Explicit
Option Compare Text
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzRfhTbSku_WithOH."

Sub RfhTbSku_WithOHxxx()
DoCmd.SetWarnings False
RunCQ "Select Distinct Sku into [#OHSkuHst] from OH"
RunCQ "Select Distinct Sku into [#OHSkuCur] from OH" & WhLasOH
RunCQ "Update Sku x Left Join [#OHSkuCur] a on x.Sku=a.Sku set WithOHCur=Not IsNull(a.Sku)"
RunCQ "Update Sku x Left Join [#OHSkuHst] a on x.Sku=a.Sku set WithOHHst=Not IsNull(a.Sku)"
RunCQ "Drop Table [#OHSkuHst]"
RunCQ "Drop Table [#OHSkuCur]"
End Sub
