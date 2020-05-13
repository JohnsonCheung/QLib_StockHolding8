Attribute VB_Name = "MxIdeSrcDclUdtUd"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcDclUdtUd."
Type UdtMbr: Mbn As String: IsAy As Boolean: Tyn As String: End Type 'Deriving(Ctor Ay)
Type Udt
    IsPrv As Boolean
    Udtn As String
    Mbr() As UdtMbr
    IsGenUdtCtor As Boolean
    IsGenUdtAy As Boolean
    IsGenUdtOpt As Boolean
    Rmk As String ' It comes from the rmk of lasLn aft rmv the Deriving(...)
End Type 'Deriving(Ay Ctor)
Function UdtMbr(IsAy As Boolean, Mbn$, Tyn$) As UdtMbr
With UdtMbr
    .IsAy = IsAy
    .Mbn = Mbn
    .Tyn = Tyn
End With
End Function


Sub PushUdtMbr(O() As UdtMbr, M As UdtMbr)
Dim N&: N = UdtMbrSi(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function Udt(IsPrv, Udtn, Mbr() As UdtMbr, IsGenUdtCtor, IsGenUdtAy, IsGenUdtOpt) As Udt
With Udt
    .IsPrv = IsPrv
    .Udtn = Udtn
    .Mbr = Mbr
    .IsGenUdtCtor = IsGenUdtCtor
    .IsGenUdtAy = IsGenUdtAy
    .IsGenUdtOpt = IsGenUdtOpt
End With
End Function

Function UdtMbrSi&(A() As UdtMbr): On Error Resume Next: UdtMbrSi = UBound(A) + 1: End Function
Function UdtMbrUB&(A() As UdtMbr): UdtMbrUB = UdtMbrSi(A) - 1: End Function

Function UdtSi&(A() As Udt): On Error Resume Next: UdtSi = UBound(A) + 1: End Function
Function UdtUB&(A() As Udt): UdtUB = UdtSi(A) - 1: End Function

Sub PushUdt(O() As Udt, M As Udt)
Dim N&: N = UdtSi(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Function SampUdt() As Udt
SampUdt = Udt(True, "UdtA", W2Mbr, True, True, True)
End Function

Private Function W2Mbr() As UdtMbr()
PushUdtMbr W2Mbr, UdtMbr(True, "MbrA", "String")
PushUdtMbr W2Mbr, UdtMbr(True, "MbrB", "String")
PushUdtMbr W2Mbr, UdtMbr(True, "MbrC", "Worksheet")
End Function

Private Sub AA_DeriUdtFun(): End Sub
Function IsGenzUdt(U As Udt) As Boolean ' Is a Udt has some code to generate
With U
Select Case True
Case .IsGenUdtAy, .IsGenUdtCtor, .IsGenUdtOpt: IsGenzUdt = True
End Select
End With
End Function

