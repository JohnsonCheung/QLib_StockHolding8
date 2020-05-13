Attribute VB_Name = "MxIdeMthOpAliAlias"
Option Explicit
Option Compare Text
Const CNs$ = "Mth.Ali.Alias"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthOpAliAlias."

Sub ACmdApply()
AliMthzN "QXls_Cmd_ApplyFilter", "CmdApply"
End Sub

Sub AU()
AliMth Upd:=eUpdAndRpt
End Sub

Sub AUO()
AliMth Upd:=eUpdOnly
End Sub
