Attribute VB_Name = "MxVbDtaNyAv"
Option Explicit
Option Compare Text
Const CNs$ = "Thw.Msg"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbDtaNyAv."
Type NyAv: Ny() As String: Av() As Variant: End Type

Function NyAv(Ny$(), Av()) As NyAv
With NyAv
    .Ny = Ny
    .Av = Av
End With
If Si(Ny) <> Si(Av) Then Raise FmtQQ("NyAv: Ny-Si[?] <> Av-Si[?]", Si(Ny), Si(Av))
End Function

Function NyAvzNav(Nav) As NyAv
If Si(Nav) = 0 Then Exit Function
NyAvzNav = NyAv(Termy(Nav(0)), CvAv(AeFstEle(Nav)))
End Function
