Attribute VB_Name = "MxIdeSrcDclUd"
Option Explicit
Option Compare Text
Const CNs$ = "Udt.MthSig"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcDclUd."
Enum eOptCpr: eCprNone: eCprTxt: eCprBin: eCprDb: End Enum
Type VbCnst: Cnstn As String: IsPrv As Boolean: TyChr As String: Tyn As String: V As String: End Type
Type VbVar: Varn As String: IsPrv As Boolean: IsAy As Boolean: TyChr As String: Tyn As String: End Type
Type VbEnmMbr: Mbn As String: Enmv As Long: End Type
Type VbEnm: Enmn As String: IsPrv As Boolean: Mbr() As VbEnmMbr: End Type
Type VbDcl
    OptExp As Boolean
    OptCpr As eOptCpr
    OptBas As Byte
    Cnst() As VbCnst
    Var() As VbVar
    Udt() As Udt
    Enm() As VbEnm
End Type
Type VbMth
    Mth As Msig
End Type
Type VbMd
    Mdn As String
    Dcl As VbDcl
    Mth() As VbMth
    LasRmk() As String
End Type
Type VbPj
    Pjn As String
    Pjf As String
    Md() As VbMd
End Type
