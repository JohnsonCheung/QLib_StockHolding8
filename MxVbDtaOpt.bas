Attribute VB_Name = "MxVbDtaOpt"
Option Explicit
Option Compare Text
Const CNs$ = "Vb.Dta"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbDtaOpt."
Type Opt
    Som As Boolean
    Itm As Variant
End Type
Sub PushIOpt(OAy, M As Opt)
If M.Som Then PushI OAy, M.Itm
End Sub

Function SomItm(Itm) As Opt
SomItm.Som = True
SomItm.Itm = Itm
End Function
