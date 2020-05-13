Attribute VB_Name = "MxDtaDaSrtDyKey"
Option Compare Text
Option Explicit
Const CNs$ = "Dy.Srt"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaSrtDyKey."
Private Type SplitDash
    SubFny() As String
    IsDes() As Boolean
End Type
Type DySrtKey
    Cxy() As Integer
    IsDes() As Boolean
End Type

Function DySrtKey(DashFF$, Fny$()) As DySrtKey
If DashFF = "" Then Exit Function
Dim DashFny$():    DashFny = SyzSS(DashFF)
With DySrtKey
    Dim A As SplitDash: A = SplitDash(DashFny)
    .Cxy = Cxy(Fny, A.SubFny)
    .IsDes = A.IsDes
End With
End Function
Function SrtgDySngKey(C&, IsDes As Boolean) As DySrtKey
With SrtgDySngKey
    .Cxy = IntAy(C)
    .IsDes = BoolAy(IsDes)
End With
End Function

Function SplitDash(DashFny$()) As SplitDash
Dim Fny$(), IsDes() As Boolean, U&
U = UB(DashFny)
ReDim Fny(U)
ReDim IsDes(U)
Fny = DashFny
Dim J%
Dim F: For Each F In DashFny
    If FstChr(F) = "-" Then
        IsDes(J) = True
        Fny(J) = RmvFstChr(F)
    End If
    J = J + 1
Next
SplitDash.IsDes = IsDes
SplitDash.SubFny = Fny
End Function
