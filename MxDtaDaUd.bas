Attribute VB_Name = "MxDtaDaUd"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDtaDaUd."

Type Drs: Fny() As String: Dy() As Variant: End Type 'Deriving(Ctor)
Type Dt: DtNm As String: Fny() As String: Dy() As Variant: End Type 'Deriving(Ctor)
Type Rec: Fny() As String: Dr() As Variant: End Type 'Deriving(Ctor)
Type Ds: DsNm As String: DtAy() As Dt: End Type ' Deriving(Ctor)


Function Drs(Fny$(), Dy() As Variant) As Drs
With Drs
    .Fny = Fny
    .Dy = Dy
End With
End Function
Function Dt(DtNm, Fny$(), Dy() As Variant) As Dt
With Dt
    .DtNm = DtNm
    .Fny = Fny
    .Dy = Dy
End With
End Function
Function Rec(Fny$(), Dr() As Variant) As Rec
With Rec
    .Fny = Fny
    .Dr = Dr
End With
End Function
Function Ds(DsNm, DtAy() As Dt) As Ds
With Ds
    .DsNm = DsNm
    .DtAy = DtAy
End With
End Function
