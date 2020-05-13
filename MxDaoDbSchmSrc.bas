Attribute VB_Name = "MxDaoDbSchmSrc"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoDbSchmSrc."
Const KwEle$ = "Ele"
Const KwEleFld$ = "EleFld"
Const KwFldDes$ = "FldDes"
Const KwKey$ = "Key"
Const KwTbl$ = "Tbl"
Const KwTblDes$ = "TblDes"
Const KwTblFldDes$ = "TblFldDes"

Sub SchmSrczS__Tst()
Dim Act As SchmSrc, Schm$(), Ept As SchmSrc
GoSub T1
Exit Sub
T1:
    Schm = SampSchm(1)
    GoTo Tst
Tst:
    Act = SchmSrczS(Schm)
    Stop
    Return
End Sub

Function SchmSrczS(Schm$()) As SchmSrc
Dim S$(): S = Schm
With SchmSrczS
    .Ele = W1Ele(W1LLny(S, KwEle))
    .EleFld = W1EleFld(W1LLny(S, KwEleFld))
    .FldDes = W1FldDes(W1LLny(S, KwFldDes))
    .Key = W1Key(W1LLny(S, KwKey))
    .Tbl = W1Tbl(W1LLny(S, KwTbl))
    .TblDes = W1TblDes(W1LLny(S, KwTblDes))
    .TblFldDes = W1TblfldDes(W1LLny(S, KwTblFldDes))
End With
End Function
Private Function W1LLny(IndtSrc$(), Key$) As LLn()
Stop
End Function
Private Function W1Ele(L() As LLn) As SmsEle()
Dim M As LLn
Dim J%: For J = 0 To LLnUB(L)
    M = L(J)
    PushSmsEle W1Ele, SmsEle(M.Lno, T1(M.Ln), RmvT1(M.Ln))
Next
End Function

Private Function W1EleFld(L() As LLn) As SmsEleFld()
Dim M As LLn
Dim J%: For J = 0 To LLnUB(L)
    M = L(J)
    PushSmsEleFld W1EleFld, SmsEleFld(M.Lno, T1(M.Ln), SyzSS(RmvT1(M.Ln)))
Next
End Function
Private Function W1FldDes(L() As LLn) As SmsFldDes()

End Function
Private Function W1Key(L() As LLn) As SmsEle()

End Function
Private Function W1Tbl(L() As LLn) As SmsTbl()
Dim M As LLn
Dim J%: For J = 0 To LLnUB(L)
    M = L(J)
    Dim Tbn$, Rst$: AsgTRst M.Ln, Tbn, Rst
    Dim Fny$(), SkFny$(): W1TblAsg Tbn, Rst, Fny, SkFny
    PushSmsTbl W1Tbl, SmsTbl(M.Lno, Tbn, Fny, SkFny)
Next
Stop
End Function
Private Sub W1TblAsg(Tbn$, Rst$, OFny$(), OSkFny$())
Dim R$: R = Replace(Rst, "*", Tbn)
Dim S As S12: S = Brk2(R, "|")
OSkFny = SyzSS(S.S1)
OFny = AddSy(OSkFny, SyzSS(S.S2))
End Sub
Private Function W1TblDes(L() As LLn) As SmsTblDes()

End Function
Private Function W1TblfldDes(L() As LLn) As SmsTblFldDes()

End Function

Private Sub SchmSrc__Tst()
Dim Act As SchmSrc, Schm$()
GoSub T1
Exit Sub
T1:
    Act = SchmSrczS(Schm)
    Stop
    Return
End Sub
