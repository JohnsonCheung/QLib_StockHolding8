Attribute VB_Name = "MxVbDtaNav"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Str.Msg"
Const CMod$ = CLib & "MxVbDtaNav."

Function AddNav(Nav1(), Nav2()) As Variant()
If Si(Nav1) = 0 Then AddNav = Nav2: Exit Function
If Si(Nav2) = 0 Then AddNav = Nav1: Exit Function
Dim O(): PushI O, Nav1(0) & " " & Nav2(0)
Dim J&
For J = 1 To UB(Nav1)
    PushI O, Nav1(J)
Next
For J = 1 To UB(Nav2)
    PushI O, Nav2(J)
Next
AddNav = O
End Function

Sub AsgNav(Nav, ONy$(), OAv())
If Si(Nav) = 0 Then Erase ONy, OAv: Exit Sub
ONy = Termy(Nav(0))
OAv = RmvFstEle(Nav)
End Sub

Function FmtNyAv(Ny$(), Av()) As String()
If Si(Ny) <> Si(Av) Then RaiseQQ "FmtNyAv: Given Ny and Av is invalid: NySi[?] AvSi[?]", Si(Ny), Si(Av)
Dim N$(): N = AmAli(Ny)
Dim J%
For J = 0 To UB(Ny)
    N(J) = Ny(J) & " " & QuoBkt(TypeName(Av(J)))
Next
N = AmAli(N)
For J = 0 To UB(N)
    PushIAy FmtNyAv, FmtNmv(N(J), Av(J))
Next
End Function

Function FmtNav(Nav) As String()
Dim J%, O$(), N$(), Av()
AsgNav Nav, N, Av
FmtNav = FmtNyAv(N, Av)
End Function

Private Sub FmtNav__Tst()
D FmtNav(Av("aa bb", 1, 2))
End Sub

Function FmtlFmsgNav$(Fun$, Msg$, Nav())
FmtlFmsgNav = FmtlFmsg(Fun, Msg) & AddPfxVbarIfNB(FmtlNav(Nav))
End Function

Function FmtFmsgNav(Fun$, Msg$, Nav()) As String()
Dim A$(): A = FmtFmsg(Fun, Msg)
Dim B$(): B = IndtSy(FmtNav(Nav))
FmtFmsgNav = AddAy(A, B)
End Function

Function FmtMsgNav(Msg$, Nav()) As String()
FmtMsgNav = AddEleAy(Msg, IndtSy(FmtNav(Nav)))
End Function

Function FmtlMsgNav$(Msg$, Nav())
FmtlMsgNav = EnsSfxDot(Msg) & " | " & FmtlNav(Nav)
End Function
