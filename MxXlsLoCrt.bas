Attribute VB_Name = "MxXlsLoCrt"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "MxXlsLoCrt."
Function CrtLoByWsDta(Ws As Worksheet) As ListObject
Set CrtLoByWsDta = CrtLo(WsDtaRg(Ws))
End Function
Function CrtLo(Rg As Range) As ListObject
Set CrtLo = WszRg(Rg).ListObjects.Add(xlSrcRange, Rg, , xlYes)
Rg.EntireColumn.AutoFit
End Function
