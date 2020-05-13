Attribute VB_Name = "MxApp"
Option Compare Text
Option Explicit
Const CNs$ = "App"
Const CLib$ = "QApp."
Const CMod$ = CLib & "MxApp."
Private Type A
    Root As String
    Nm As String
    Ver As String
    OupHom As String
End Type
Private A As A

Function TpFx$(Optional Tpn$)

End Function

Function TpFxm$(Optional Tpn$)
TpFxm = AppHom & TpFxmFn(DftTpn(Tpn))
End Function

Private Function DftTpn$(Tpn$)
DftTpn = IIf(Tpn = "", A.Nm, Tpn)
End Function

Function AppPth$()
On Error GoTo E
Static O$
With A
If O = "" Then O = AddFdrAp(.Root, .Nm, .Ver)
End With
AppPth = O
E:
End Function


Function TpFxFn$(Tpn$)
TpFxFn = Tpn & "(Template).xlsx"
End Function

Function TpFxmFn$(Tpn$)
TpFxmFn = Tpn & "(Template).xlsm"
End Function

Function OupPth$()
On Error GoTo E
OupPth = AddFdrAp(A.OupHom, A.Nm)
E:
End Function

Property Get OupFxzNxt$()
On Error GoTo E
OupFxzNxt = NxtFfnzAva(OupFx)
E:
End Property

Property Get OupFx$()
On Error GoTo E
OupFx = OupPth & A.Nm & ".xlsx"
E:
End Property
'
Property Get AppFb$()
On Error GoTo E
'C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxExpCmp\TaxExpCmp\TaxExpCmp.1_3..accdb
AppFb = AppPth & "AppFb.accdb"
E:
End Property
'
Function AppHom$()
On Error GoTo E
Static Y$
With A
If Y = "" Then Y = AddFdrApEns(.Root, .Nm, .Ver)
End With
AppHom = Y
E:
End Function
