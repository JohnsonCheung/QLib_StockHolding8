Attribute VB_Name = "MxVbStrBktNm"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbStrBktNm."

Private Sub RmvNmBkt__Tst()
Dim S$, Nm$, OpnBkt$
GoSub T1
Exit Sub
T1:
    S = "AA1 B(A())x"
    Ept = "AA1 x"
    OpnBkt = "("
    GoTo Tst
T2:
    S = "aaa B(lsdfj)1aa"
    Nm = "B"
    Ept = "aaa 1aa"
    OpnBkt = "("
    GoTo Tst
Tst:
    Act = RmvNmBkt(S, Nm, OpnBkt)
    C
    Return
End Sub

Function NmBktStr$(S, Nm$, Optional OpnBkt$ = vbOpnBkt)
NmBktStr = BetP12(S, XStrP12(S, Nm, OpnBkt))
End Function

Function RmvNmBkt$(S, Nm$, Optional OpnBkt$ = vbOpnBkt) 'Ret a str aft rmv the NmBktStr, which is a of Nm(...)
RmvNmBkt = RmvP12(S, XStrP12(S, Nm, OpnBkt))
End Function

Function BetNmBkt$(S, Nm$, Optional OpnBkt$ = vbOpnBkt) 'Ret str bet a NmBkt
BetNmBkt = BetP12(S, XBktP12(S, Nm, OpnBkt))
End Function

'==X
Private Sub XOpnPos__Tst()
Dim S, Nm, OpnBkt$
GoSub T1
Exit Sub
T1:
    S = 1
    Nm = 1
    OpnBkt = 1
    GoTo Tst
Tst:
    Act = XOpnPos(S, Nm, OpnBkt)
    C
    Return
End Sub

Private Function XOpnPos%(S, Nm, Optional OpnBkt$ = vbOpnBkt) ' ret pos of zOpn, which is the OpnBkt of the NmBkt
Dim B%: B = XNmPos%(S, Nm, OpnBkt): If B = 0 Then Exit Function
XOpnPos = B + Len(Nm)
End Function

Private Function XNmPos%(S, Nm, Optional OpnBkt$ = vbOpnBkt) ' ret pos of Nm, which the Nm of the NmBkt
Dim Opn$
Opn = Nm & OpnBkt: If HasPfx(S, Opn) Then XNmPos = 1: Exit Function
Opn = " " & Opn
Dim P%: P = InStr(S, Opn): If P = 0 Then Exit Function
XNmPos = P - 1
End Function

Private Function XStrP12(S, Nm, Optional OpnBkt$ = vbOpnBkt) As C12 ' NmBktStr is the str of Nm(...)
Dim C As C12: C = XBktP12(S, Nm, OpnBkt): If IsEmpC12(C) Then Exit Function
XStrP12 = C12(C.C1 - Len(Nm), C.C2)
End Function

Private Function XBktP12(S, Nm, Optional OpnBkt$ = vbOpnBkt) As C12 ' ret BktP12 of the NmBkt
Dim Opn%: Opn = XOpnPos(S, Nm, OpnBkt): If Opn = 0 Then Exit Function
XBktP12 = C12(Opn, ClsBktPos(S, Opn, OpnBkt))
End Function
