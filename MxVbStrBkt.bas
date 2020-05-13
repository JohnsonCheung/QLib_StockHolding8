Attribute VB_Name = "MxVbStrBkt"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrBkt."

Private Sub BktP12__Tst()
Dim A$, Act As C12, Ept As C12
'
A = "(A(B)A)A"
Ept = C12(1, 7)
GoSub Tst
'
A = " (A(B)A)A"
Ept = C12(2, 8)
GoSub Tst
'
A = " (A(B)A )A"
Ept = C12(2, 9)
GoSub Tst
'
Exit Sub
Tst:
    Act = BktP12(A)
    Debug.Assert IsEqC12(Act, Ept)
    Return
End Sub

Function BktS12(OpnBkt$) As S12
With BktS12
    .S1 = OpnBkt
    .S2 = ClsBkt(OpnBkt)
End With
End Function

Function BktP12(S, Optional OpnBkt$ = vbOpnBkt) As C12
Const CSub$ = CMod & "BktP12"
Dim OpnPos%: OpnPos = InStr(S, OpnBkt): If OpnPos = 0 Then Exit Function
Dim ClsPos%: ClsPos = ClsBktPos(S, OpnPos, OpnBkt)
BktP12 = C12(OpnPos, ClsPos)
End Function

Function ClsBktPos%(S, OpnBktPos%, OpnBkt$)
If Mid(S, OpnBktPos, 1) <> OpnBkt Then PmEr CSub, "@OpnBktPos not point to a @OpnBkt", "Chr-At-OpnBktPos OpnBkt OpnBktPos S", Mid(S, OpnBktPos, 1), OpnBkt, OpnBktPos, S
Dim Cls$: Cls = ClsBkt(OpnBkt)
Dim NOpn%, O%
Dim J%: For J = OpnBktPos + 1 To Len(S)
    Select Case Mid(S, J, 1)
    Case Cls
        If NOpn = 0 Then
            O = J
            ClsBktPos = O
            Exit Function
        End If
        NOpn = NOpn - 1
    Case OpnBkt
        NOpn = NOpn + 1
    End Select
Next
Thw CSub, "There is OpnBkt in S and ClsBkt is missing", "OpnBkt ClsBkt S OpnBktPos", OpnBkt, Cls, S, OpnBktPos
End Function

Function ClsBkt$(OpnBkt$)
Select Case OpnBkt
Case "(": ClsBkt = ")"
Case "[": ClsBkt = "]"
Case "{": ClsBkt = "}"
Case Else: Stop
End Select
End Function

Private Sub BrkBkt123__Tst()
Dim A$, OpnBkt$
A = "aaaa((a),(b))xxx":    OpnBkt = "(":          Ept = Sy("aaaa", "(a),(b)", "xxx"): GoSub Tst
Exit Sub
Tst:
    Act = BrkBkt123(A, OpnBkt)
    C
    Return
End Sub

Function BrkBkt123(S, Optional OpnBkt$ = vbOpnBkt) As String() ' Ret 3 string as Sy which is (Bef Bet Aft)-Bkt
Dim P As C12: P = BktP12(S, OpnBkt): If IsEmpC12(P) Then Exit Function
BrkBkt123 = Sy( _
    Left(S, P.C1 - 1), _
    BetPos(S, P), _
    Mid(S, P.C2 + 1))
End Function

Function BetPos$(S, Pos As C12)
If IsEmpC12(Pos) Then Exit Function
BetPos = Bet(S, Pos.C1, Pos.C2)
End Function

Function BetBktMust$(S, Fun$, Optional OpnBkt$ = vbOpnBkt)
BetBktMust = BetBkt(S, OpnBkt)
If BetBktMust = "" Then Thw Fun, "No Bkt is found in Str", "Str", S
End Function

Function BetBkt$(S, Optional OpnBkt$ = vbOpnBkt)
BetBkt = BetPos(S, BktP12(S, OpnBkt))
End Function

Function AftBkt$(S, Optional OpnBkt$ = vbOpnBkt)
Dim P As C12: P = BktP12(S, OpnBkt): If IsEmpC12(P) Then Exit Function
AftBkt = Mid(S, P.C2 + 1)
End Function

Function BefBkt$(S, Optional OpnBkt$ = vbOpnBkt)
Dim P As C12: P = BktP12(S, OpnBkt): If IsEmpC12(P) Then Exit Function
BefBkt = Left(S, P.C1 - 1)
End Function
