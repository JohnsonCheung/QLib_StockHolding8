Attribute VB_Name = "gzStkHldPm"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "gzStkHldPm."
Function MB52IPthPm$():          MB52IPthPm = CPmv("MB52_InpPth"):      End Function
Function IsCpy1Pm() As Boolean:    IsCpy1Pm = CPmv("MB52_IsCpyToPth1"): End Function
Function IsCpy2Pm() As Boolean:    IsCpy2Pm = CPmv("MB52_IsCpyToPth2"): End Function
Function CpyToPth2Pm$():        CpyToPth2Pm = CPmv("MB52_CpyToPth2"): End Function
Function CpyToPth1Pm$():        CpyToPth1Pm = CPmv("MB52_CpyToPth1"): End Function
Function SkuIFxPm$():              SkuIFxPm = CPmv("Sku_InpFx"): End Function
