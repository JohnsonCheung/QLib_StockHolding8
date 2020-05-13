Attribute VB_Name = "gzSku__Fun"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzSku__Fun."
Function SkuPrm(Pmn$):    SkuPrm = CPmv("Sku_" & Pmn):     End Function
Function SkuPrm_InpFx$(): SkuPrm_InpFx = SkuPrm("InpFx"): End Function
Function SkuLisPrm_CpyToPth$():          SkuLisPrm_CpyToPth = CPmv("SkuLis_CpyToPth"): End Function
Function SkuLisPrm_IsCpyTo() As Boolean: SkuLisPrm_IsCpyTo = CPmv("SkuLis_IsCpyTo"):   End Function
