Attribute VB_Name = "MxTpSqSw"
Option Explicit
Option Compare Text
Private Type XPmSw: Pm As Dictionary: Sw As Dictionary: End Type 'Deriving(Ctor Ay)
Type SqSw: Pm As Dictionary: StmtSw As Dictionary: FldSw As Dictionary: End Type
Type SwOral: A As String: End Type
Type SwEqnl: A As String: End Type
Enum eSwOp: eOraSw: eEqnSw: End Enum
Function SampSqSw() As SqSw
With SampSqSw
Set .Pm = SampPm
Set .FldSw = SampFldSw
Set .StmtSw = SampStmtSw
End With
End Function
Private Function SampFldSw() As Dictionary

End Function
Private Function SampStmtSw() As Dictionary

End Function
Function SampPm() As Dictionary
Set SampPm = Dic(SampPmSrc)
End Function
Function SampPmSrc() As String()

End Function
Function SampSwSrc() As String()
Erase XX
X "?LvlY    EQ >>SumLvl Y"
X "?LvlM    EQ >>SumLvl M"
X "?LvlW    EQ >>SumLvl W"
X "?LvlD    EQ >>SumLvl D"
X "?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY"
X "?M       OR ?LvlD ?LvlW ?LvlM"
X "?W       OR ?LvlD ?LvlW"
X "?D       OR ?LvlD"
X "?Dte     OR ?LvlD"
X "?Mbr     OR >?BrkMbr"
X "?MbrCnt  OR >?BrkMbr"
X "?Div     OR >?BrkDiv"
X "?Sto     OR >?BrkSto"
X "?Crd     OR >?BrkCrd"
X "?#SEL#Div NE >LisDiv *blank"
X "?#SEL#Sto NE >LisSto *blank"
X "?#SEL#Crd NE >LisCrd *blank"
SampSwSrc = XX
Erase XX
End Function
'-------

Function SqSwzS() As SqSw
End Function

Private Function SqSw(L() As Swl, P As Dictionary) As SqSw
End Function

'Pub ======================================================================
Function SqSwzLy(SwLy$(), Pm As Dictionary) As SqSw
SqSwzLy = UUEvl(W1Swly(SwLy), Pm)
End Function

Private Function UUEvl(L() As Swl, Pm As Dictionary) As SqSw
Dim OSw As New Dictionary
Dim J%
Do
    If Not UUEvlR(L, OSw, Pm) Then W2Thw L, OSw, Pm
    If SwlSi(L) = 0 Then UUEvl = W2Sw2(OSw): Exit Function
    LoopTooMuch CSub, J
Loop
End Function
Private Function W1Swly(SwLy$()) As Swl()
Dim L: For Each L In Itr(SwLy)
    PushSwl W1Swly, W1Swl(L)
Next
End Function
Private Function W1Swl(Ln) As Swl
Const CSub$ = CMod & "SwlzLin"
Dim Ay$(): Ay = Termy(Ln)
Dim Nm$: Nm = Ay(0)
Dim OpStr$: OpStr = Ay(1)
Dim Op As eBoolOp: Op = eBoolOp(OpStr)
Dim T1$, T2$
Select Case True
Case Op = eOpNe, Op = eOpEq:
    If Si(Ay) <> 4 Then Thw CSub, "Ln should have 4 terms for Eq | Ne", "Ln", Ln
    T1 = Ay(2): T2 = Ay(3):
Case Op = eOpAnd, Op = eOpOr
    If Si(Ay) < 3 Then Thw CSub, "Ln should have at 3 terms And | Or", "Ln", Ln
    Ay = AeFstNEle(Ay, 2)
End Select
W1Swl = Swl(Nm, Op, Ay)
End Function

Private Sub W2Thw(L() As Swl, Sw As Dictionary, Pm As Dictionary)
Dim Left$(), Evld$()
    Left = FmtSwly(L)
    Evld = FmtDic(Sw)
Thw "Sw2zLy", "Cannot eval all Swl", "Some Swl is left", "SwLy Pm [Swl left] [evaluated Swl]", FmtSwly(L), Pm, Left, Evld
End Sub
Private Function W2Sw2(Sw As Dictionary) As SqSw
Dim Fld As New Dictionary
Dim Stmt As New Dictionary
Dim K: For Each K In Sw.Keys
    If HasPfx(K, "?") Then
        Stmt.Add RmvFstChr(K), Sw(K)
    Else
        Fld.Add K, Sw(K)
    End If
Next
Set W2Sw2.FldSw = Fld
Set W2Sw2.StmtSw = Stmt
End Function

Private Function UUEvlR(O() As Swl, OSw As Dictionary, Pm As Dictionary) As Boolean
Dim M As Swl, Ixy%(), OHasEvl As Boolean
Dim J%: For J = 0 To SwlUB(O)
    M = O(J)
    With UUEvlLn(M, XPmSw(Pm, OSw))
        If .Som Then
            OSw.Add M.Swn, .Bool
            OHasEvl = True
            PushI Ixy, J
        End If
    End With
Next
O = W3RmvSwl(O, Ixy)
UUEvlR = OHasEvl
End Function
Private Function W3RmvSwl(L() As Swl, Ixy%()) As Swl()
Dim J%: For J = 0 To SwlUB(L)
    If Not HasEle(Ixy, J) Then PushSwl W3RmvSwl, L(J)
Next
End Function
Private Function UUEvlLn(L As Swl, P As XPmSw) As BoolOpt
With L
Select Case True
Case IsOrAndStr(.Op): UUEvlLn = UUEvlAndOr(CvOrAndStr(.Op), .Termy, P)
Case IsEqNeStr(.Op):  UUEvlLn = UUEvlEqNe(CvEqNeStr(.Op), .Termy(0), .Termy(1), P)
End Select
End With
End Function
Private Function UUEvlAndOr(Op As eOrAnd, Termy$(), P As XPmSw) As BoolOpt
Dim O() As Boolean
Dim T: For Each T In Itr(Termy)
    With W5EvlTerm(T, P)
        If Not .Som Then Exit Function
        PushI O, .Bool
    End With
Next
UUEvlAndOr = EvlBooly(O, Op)
End Function
Private Function W5EvlTerm(T, P As XPmSw) As BoolOpt
Select Case True
Case P.Sw.Exists(T): W5EvlTerm = SomBool(P.Sw(T))
Case P.Pm.Exists(T): W5EvlTerm = SomBool(P.Pm(T))
End Select
End Function

Private Function UUEvlEqNe(Op As eEqNe, T1$, T2$, P As XPmSw) As BoolOpt 'Return True and set ORslt if evaluated
Dim S1$
    With W6EvlT1(T1, P.Pm)
        If Not .Som Then Exit Function
        S1 = .Str
    End With
Dim S2$
    With W6EvlT2(T2, P.Pm)
        If Not .Som Then Exit Function
        S2 = .Str
    End With
Select Case True
Case Op = eEqnEq: UUEvlEqNe = SomBool(S1 = S2)
Case Op = eEqnNe: UUEvlEqNe = SomBool(S1 <> S2)
Case Else: Imposs CSub
End Select
End Function
Private Function W6EvlT1(T1$, Pm As Dictionary) As StrOpt
If Pm.Exists(T1) Then W6EvlT1 = SomStr(Pm(T1))
End Function
Private Function W6EvlT2(T2$, Pm As Dictionary) As StrOpt
If T2 = "*Blank" Then W6EvlT2 = SomStr(""): Exit Function
Dim M As StrOpt: M = W6EvlT1(T2, Pm)
If M.Som Then W6EvlT2 = M: Exit Function
W6EvlT2 = SomStr(T2)
End Function

Private Function XPmSw(Pm As Dictionary, Sw As Dictionary) As XPmSw
With XPmSw
    Set .Pm = Pm
    Set .Sw = Sw
End With
End Function
Function AddXPmSw(A As XPmSw, B As XPmSw) As XPmSw(): PushXPmSw AddXPmSw, A: PushXPmSw AddXPmSw, B: End Function
Sub PushXPmSwAy(O() As XPmSw, A() As XPmSw): Dim J&: For J = 0 To XPmSwUB(A): PushXPmSw O, A(J): Next: End Sub
Sub PushXPmSw(O() As XPmSw, M As XPmSw): Dim N&: N = XPmSwSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function XPmSwSi&(A() As XPmSw): On Error Resume Next: XPmSwSi = UBound(A) + 1: End Function
Function XPmSwUB&(A() As XPmSw): XPmSwUB = XPmSwSi(A) - 1: End Function
