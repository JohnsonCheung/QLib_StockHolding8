Attribute VB_Name = "MxIdeSrcDclItm"
Option Compare Text
Option Explicit
Const CNs$ = "Ide.Dcl"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcDclItm."

Const C_Enm$ = "Enum"
Const C_Ty$ = "Type"

':ShtDclSfx: :Sfx
':DclSfx:    :Sfx
':Dcn:   :Nm #Dcl-Itm-Nm#
':Dcitm: :S  #Dcl-Itm# ! Is from [Dim] | [Arg]-aft-rmv-[=]-Optional-Paramarray | End-Type
Function Dcn$(Dcitm)
If HasSubStr(Dcitm, " As ") Then
    Dcn = DcnoAs(Dcitm)
Else
    Dcn = DcnoTyChr(Dcitm)
End If
End Function

Function DcnoTyChr$(DimShtItm)
DcnoTyChr = RmvLasChrzLis(RmvSfxzBkt(DimShtItm), TyChrLis)
End Function

Function DcnoAs$(DimAsItm)
DcnoAs = RmvSfxzBkt(Bef(DimAsItm, " As"))
End Function

Function DclSfx$(Dcitm$)
DclSfx = ShtDclSfx(RmvNm(Dcitm))
End Function

Function ShtDclSfx$(DclSfx$)
If DclSfx = "" Then Exit Function
Dim L$: L = DclSfx
Select Case True
Case L = " As Boolean":: ShtDclSfx = "^"
Case L = " As Boolean()": ShtDclSfx = "^()"
Case Else
    ShfPfx L, " As "
    ShtDclSfx = L
End Select
End Function
