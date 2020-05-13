Attribute VB_Name = "MxVbDtaOupOpt"
Option Explicit
Option Compare Text
Const CNs$ = "OupOpt"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxVbDtaOupOpt."
Enum eOupTy: eDmpOup: eBrwOup: eVcOup: End Enum
Type OupOpt: FnPfx As String: OupTy As eOupTy: End Type
Function OupOpt(Optional FnPfx$, Optional OupTy As eOupTy = eOupTy.eBrwOup) As OupOpt
With OupOpt
    .FnPfx = FnPfx
    .OupTy = OupTy
End With
End Function

Function Dmpg() As OupOpt: Dmpg = OupOpt(, eDmpOup): End Function
Function Vcg(Optional FnPfx$) As OupOpt: Vcg = OupOpt(FnPfx, eVcOup): End Function
Function Brwg(Optional FnPfx$) As OupOpt: Brwg = OupOpt(FnPfx, eBrwOup): End Function
