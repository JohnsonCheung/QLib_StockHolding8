Attribute VB_Name = "MxVbRunThw"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Thw"
Const CMod$ = CLib & "MxVbRunThw."
Enum eHaltTy: ePgmEr: ePmEr: eLgcEr: eImposs: eLoopTooMuch: eUsrInf: eUsrWarn: End Enum
Public Const eHaltTySS$ = "ePgmEr ePmEr eLgcEr eImposs eLoopTooMuch eUsrInf eUsrWarn"
Const eHaltTyTxt0$ = "Program Error"
Const eHaltTyTxt1$ = "Logic Error"
Const eHaltTyTxt2$ = "Impossible to reach here"
Const eHaltTyTxt3$ = "Looping too much"
Const eHaltTyTxt4$ = "User information"
Const eHaltTyTxt5$ = "User warning"

Type CfgInf
    ShwInf As Boolean
    ShwTim As Boolean
End Type
Type CfgSql
    FmtSql As Boolean
End Type
Type Cfg
    Inf As CfgInf
    Sql As CfgSql
End Type

Public Property Get Cfg() As Cfg
Static X As Boolean, Y As Cfg
If Not X Then
    X = True
    Y.Sql.FmtSql = True
    Y.Inf.ShwInf = True
    Y.Inf.ShwTim = True
End If
Cfg = Y
End Property

Private Sub Thw__Tst()
Thw "SF", "AF"
End Sub

Sub InfLn(Fun$, Msg$, ParamArray Nap())
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
D FmtlFmsgNav(Fun, Msg, Nav)
End Sub

Sub WarnLn(Fun$, Msg$, ParamArray Nap())
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
Debug.Print FmtlFmsgNav(Fun, Msg, Nav)
End Sub

Sub Warn(Fun$, Msg$, ParamArray Nap())
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
D FmtFmsgNav(Fun, Msg, Nav)
End Sub

Sub LoopTooMuch(Fun$, OCnt%, Optional Max% = 10000)
OCnt = OCnt + 1: If OCnt > Max Then ThwMsg HaltTxt("Logic Error", "Looping too much", Fun)
End Sub

Private Function HaltTxt$(T As eHaltTy, Optional Fun$, Optional Msgln$, Optional Nav)
Dim Nav1()
If Msgln <> "" Then Nav1 = AddNav(Av("Msg", Msgln), CvAv(Nav))
If Fun <> "" Then Nav1 = AddNav(Av("Fun", Fun), Nav1)
HaltTxt = HaltBoxl(T) & FmtlNav(Nav1)
End Function

Function eHaltTyStr$(T As eHaltTy): eHaltTyStr = SyzSS(eHaltTySS)(T): End Function
Function eHaltTyTxt$(T As eHaltTy): eHaltTyTxt = eHaltTyTxty()(T): End Function

Private Function eHaltTyTxty() As String()
Dim X As Boolean, Y$()
If Not X Then
    X = True
    PushI Y, eHaltTyTxt0
    PushI Y, eHaltTyTxt1
    PushI Y, eHaltTyTxt2
    PushI Y, eHaltTyTxt3
    PushI Y, eHaltTyTxt4
    PushI Y, eHaltTyTxt5
End If
eHaltTyTxty = Y
End Function

Private Function HaltBoxl$(T As eHaltTy)
HaltBoxl = Boxl(eHaltTyTxt(T))
End Function

Sub RaiseChkNotePad(Optional Fun$)
Dim A$: If Fun <> "" Then A = " in subr-[" & Fun & "]"
Raise FmtQQ("There is error?.  See the message in the NotePad.", A)
End Sub

Sub ThwMsg(Msg$)
BrwStr Msg
RaiseChkNotePad
End Sub

Sub RaiseQQ(MsgQQ$, ParamArray Ap())
Dim Av(): Av = Ap
Err.Raise 1, , FmtQQAv(MsgQQ, Av)
End Sub

Sub Raise(Msg$): Err.Raise 1, , Msg: End Sub

Sub ChkEr(Er$(), Optional Fun$)
If Si(Er) = 0 Then Exit Sub
BrwAy Er, "Er_"
RaiseChkNotePad Fun
End Sub
Sub ChkErl(Erl$, Optional Fun$): ChkEr SplitCrLf(Erl), Fun: End Sub

Sub ThwTrue(IfTrue As Boolean, Fun$, Msg$, ParamArray Nap())
If IfTrue Then
Dim Av(): Av = Nap: ThwNav Fun, Msg, Av
End If
End Sub
Sub ThwNav(Fun$, Msg$, Nav())
Static IsInThw As Boolean
If IsInThw Then Raise "Thw is called recurively....."
IsInThw = True
ThwMsg HaltTxt(ePgmEr, Fun, Msg, Nav)
IsInThw = False
End Sub
Sub Thw(Fun$, Msg$, ParamArray Nap())
Dim Av(): Av = Nap: ThwNav Fun, Msg, Av
End Sub

Sub ThwNever(Fun$, Optional Msg$)
ThwMsg HaltTxt(eImposs, Fun, Msg)
End Sub

Sub PmEr(Fun$, Msgln$, ParamArray Nap())
Dim Nav(): Nav = Nap
ThwMsg HaltTxt(ePmEr, Fun, Msgln, Nav)
End Sub

Sub LgcEr(Fun$, Msgln$, ParamArray Nap())
Dim Nav(): Nav = Nap
ThwMsg HaltTxt(eLgcEr, Fun, Msgln, Nav)
End Sub

Sub EmEr(Fun$, EmNm$, Emv, VdtEmXxxSS$)
Dim S%: S = Si(SyzSS(VdtEmXxxSS))
ThwMsg HaltTxt("Enumerator invalid value error", , Fun, Array("EmNm Emv Enumerate-Val-Cnt VdtEmXxxSS", EmNm, Emv, S, VdtEmXxxSS))
End Sub

Sub Imposs(Fun$, Optional Reason$)
Thw Fun, "Impossible to reach here", Reason
End Sub

Sub Inf(Fun$, Msg$, ParamArray Nap())
If Not Cfg.Inf.ShwInf Then Exit Sub
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
D FmtFmsgNav(Fun, Msg, Nav)
End Sub
Sub EnmEr(Fun$, EmNm$, Emss$, UnexpEi&)
Dim O$()
PushI O, "Unexpected Enum Value:"
PushI O, "====================="
PushI O, "Fun    : [" & Fun & "]"
PushI O, "UnexpEi: [" & UnexpEi & "]"
PushI O, "EmNm   : [" & EmNm & "]"
PushI O, "Emss   : [" & Emss & "]"
PushI O, "NEm-Itm: [" & CntzSS(Emss) & "]"
PushI O, "UnexpEi: [" & UnexpEi & "]"
BrwAy O
Raise "Unexpected parameter.....!"

End Sub

Sub ChkIsStrOrSy(V)
Select Case True
Case IsStr(V), IsSy(V): Exit Sub
End Select
Err
End Sub
Sub BrwEr(Er_Of_Str_or_Sy)
Dim Er: Er = Er_Of_Str_or_Sy
ChkIsStrOrSy Er
If Si(Er) = 0 Then Exit Sub
BrwAy Er
End Sub
