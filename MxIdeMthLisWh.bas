Attribute VB_Name = "MxIdeMthLisWh"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Mth.Lis"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthLisWh."
Type WhMth
Patn As String: ShtMdySS As String: ShtMthTySS As String
TyChr As String: RetAsPatn As String: RetAy As eTri
FstPmTyPatn As String: NPm As Integer: ShtPmPatn As String: AnyAp As eTri
MdnPatn As String: ShtMdTySS As String
End Type

Function WhAllMth() As WhMth: End Function

Function WhMth(Optional Patn$, Optional ShtMdySS$, Optional ShtMthTySS$, _
Optional TyChr$, Optional RetAsPatn$, Optional RetAy As eTri, _
Optional FstPmTyPatn$, Optional NPm% = -1, Optional ShtPmPatn$, Optional AnyAp As eTri, _
Optional MdnPatn$, Optional ShtMdTySS$) As WhMth
With WhMth
    .Patn = Patn: .ShtMdySS = ShtMdySS: .ShtMthTySS = ShtMthTySS
    .TyChr = TyChr: .RetAsPatn = RetAsPatn: .RetAy = RetAy
    .FstPmTyPatn = FstPmTyPatn: .NPm = NPm: .ShtPmPatn = ShtPmPatn: .AnyAp = AnyAp
    .MdnPatn = MdnPatn: .ShtMdTySS = ShtMdTySS
End With
End Function
