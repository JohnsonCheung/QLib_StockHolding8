Attribute VB_Name = "MxVbDtaNavFmt"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Str.Msg"
Const CMod$ = CLib & "MxVbDtaNavFmt."
#Const Doc = False
#If Doc Then
ksdf
sdfs sdf
#End If
Function AddNNAv(Nav(), NN$, Av()) As Variant()
Dim O(): O = Nav
If Si(O) = 0 Then
    PushI O, NN
Else
    O(0) = O(0) & " " & NN
End If
PushAy O, Av
AddNNAv = O
End Function

Function AddNmV(Nav(), Nm$, V) As Variant()
AddNmV = AddNNAv(Nav, Nm, Av(V))
End Function

Sub DmpNap(ParamArray Nap())
Dim Nav(): If UBound(Nap) >= 0 Then Nav = Nap
DmpNav Nav
End Sub

Sub DmpNav(Nav())
Dmp FmtNav(Nav)
End Sub

Function FmtFmsg(Fun$, Msg$) As String()
Dim O$(), MsgL1$, MsgRst$
AsgBrk1Dot Msg, MsgL1, MsgRst
PushI FmtFmsg, EnsSfxDot(MsgL1) & IIf(Fun = "", "", "  @" & Fun)
PushIAy FmtFmsg, IndtSy(WrpLy(SplitCrLf(MsgRst)))
End Function


Function FmtFmsgNap(Fun$, Msg$, ParamArray Nap()) As String()
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
If Fun = "" And Msg = "" And Si(Nav) = 0 Then Exit Function
FmtFmsgNap = FmtFmsgNav(Fun, Msg, Nav)
End Function

Function FmtFmsgObjPP(Fun$, Msg$, Obj As Object, PP$) As String()
FmtFmsgObjPP = AddAy(FmtFmsg(Fun, Msg), LyzObjPP(Obj, PP))
End Function

Function FmtMsgNap(Msg$, ParamArray Nap()) As String()
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
FmtMsgNap = FmtMsgNav(Msg, Nav)
End Function

Function NmvzDrs(Nm$, A As Drs) As String()
NmvzDrs = FmtNmv(Nm, FmtDrs(A))
End Function

Function NmvzDrsO(Nm$, A As Drs, Opt As DrsFmto) As String()
NmvzDrsO = FmtNmv(Nm, FmtDrsNO(A, Opt))
End Function

Function FmtNmvWiTyn(Nm$, V) As String()
Dim N$: N = Nm & " (" & TypeName(V) & ")"
FmtNmvWiTyn = FmtNmv(N, V)
End Function

Function FmtNmv(Nm$, V) As String()
Dim Ly$(): Ly = FmtV(V)
PushI FmtNmv, Nm & ": " & QuoSq(FstEle(Ly))
If Si(Ly) <= 1 Then Exit Function
Dim S$: S = Space(Len(Nm) + 2)
Dim J%: For J = 1 To UB(Ly)
    PushI FmtNmv, S & QuoSq(Ly(J))
Next
End Function

Private Function TySfx$(V)
TySfx = " ::" & TypeName(V)
End Function

Function FmtlFmsg$(Fun$, Msg$)
Dim F$: F = IIf(Fun = "", "", " (@" & Fun & ")")
Dim A$: A = Msg & F
If Cfg.Inf.ShwTim Then
    FmtlFmsg = NowStr & " | " & A
Else
    FmtlFmsg = A
End If
End Function

Function FmtlNav$(Nav)
FmtlNav = JnCrLf(FmtNav(Nav))
End Function
