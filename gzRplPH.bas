Attribute VB_Name = "gzRplPH"
Option Compare Text
Option Explicit
Const CLib$ = "QAppMB52."
Const CMod$ = CLib & "gzRplPH."
Function RplL4$(QStr$):  RplL4 = RplQ(QStr, "L4"): End Function
Function RplL3$(QStr$):  RplL3 = RplQ(QStr, "L3"): End Function
Function RplL2$(QStr$):  RplL2 = RplQ(QStr, "L2"): End Function
Function RplL1$(QStr$):  RplL1 = RplQ(QStr, "L1"): End Function
Function RplSku$(QStr$): RplSku = RplQ(QStr, "Sku"): End Function
Function RplStm$(QStr$): RplStm = RplQ(QStr, "Stm"): End Function
Function RplBus$(QStr$): RplBus = RplQ(QStr, "Bus"): End Function
