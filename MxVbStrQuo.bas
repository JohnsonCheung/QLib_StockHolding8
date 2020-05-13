Attribute VB_Name = "MxVbStrQuo"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str.Quo"
Const CMod$ = CLib & "MxVbStrQuo."
#If Doc Then
'Bkt:Cml #Bracket#
'Quo:Cml #Quote#
#End If

Function BrkQuo(QuoStr$) As S12
Dim L%: L = Len(QuoStr)
Dim S1$, S2$
Select Case L
Case 0:
Case 1
    S1 = QuoStr
    S2 = QuoStr
Case 2
    S1 = Left(QuoStr, 1)
    S2 = Right(QuoStr, 1)
Case Else
    If InStr(QuoStr, "*") > 0 Then
        BrkQuo = Brk(QuoStr, "*", NoTrim:=True)
        Exit Function
    End If
    Stop
End Select
BrkQuo = S12(S1, S2)
End Function

Function QuoAy(Ay, QuoStr$) As String()
Dim P$, S$
With BrkQuo(QuoStr)
    P = .S1
    S = .S2
End With
QuoAy = AmAddPfxSfx(Ay, P, S)
End Function

Function UnQuoVb$(QuoVb)
UnQuoVb = Replace(RmvFstLasChr(QuoVb), vb2DblQ, vbDblQ)
End Function

Function QuoVb$(S)
':QuoVb: #Quoted-Vb-Str# ! a str with fst and lst chr is vbDblQ and inside each vbDblQ is in pair, which will cv to one vbDblQ  @@
QuoVb = vbDblQ & Replace(S, vbDblQ, vb2DblQ) & vbDblQ
End Function

Function Quo$(S, QuoStr$)
With BrkQuo(QuoStr)
    Quo = .S1 & S & .S2
End With
End Function

Function QuoIfNB$(S_IfNB, QuoStr$)
If Trim(S_IfNB) = "" Then Exit Function
With BrkQuo(QuoStr)
    QuoIfNB = .S1 & S_IfNB & .S2
End With
End Function

Function QuoSqlStr$(SqlStr) ' If @SqlStr has only single-quote, quote by single.  If has only double, quote single.  Else, quote single and replace inside-single as 2-single.
Dim WiSng As Boolean, WiDbl As Boolean
WiSng = HasSngQ(SqlStr)
WiDbl = HasDblQ(SqlStr)
Dim O$
Select Case True
Case WiSng And WiDbl: O = QuoSng(Replace(SqlStr, vbSngQ, vb2SngQ))
Case WiSng: O = QuoDbl(SqlStr)
Case WiDbl: O = QuoSng(SqlStr)
Case Else: Imposs CSub, "Select-case in pgm should never reach case-else here"
End Select
QuoSqlStr = O
End Function

Function QuoSqlvy(Vy) As String()
Dim V: For Each V In Vy
    PushI QuoSqlvy, QuoSqlv(V)
Next
End Function
Function QuoSqlPrimy(Primy) As String()
ChkIsPrimy Primy, CSub
Dim I
Select Case True
Case IsSy(Primy)
    For Each I In Primy
        PushI QuoSqlPrimy, QuoSqlStr(I)
    Next
Case IsDtey(Primy)
    For Each I In Primy
        PushI QuoSqlPrimy, QuoDte(I)
    Next
Case Else
    For Each I In Primy
        PushI QuoSqlPrimy, I
    Next
End Select
End Function
Function QuoSqlv$(Sqlv)
Dim V: V = Sqlv
Select Case True
Case IsStr(V): QuoSqlv = QuoSqlStr(V)
Case IsDte(V): QuoSqlv = QuoDte(V)
Case IsNum(V), IsBool(V): QuoSqlv = V
Case IsEmpty(V): QuoSqlv = "null"
Case Else
    Thw CSub, "V should be Dte Str, Numeric or Empty", "TypeName(V)", TypeName(V)
End Select
End Function

Function QuoVbStr$(S): QuoVbStr = vbDblQ & Replace(S, vbDblQ, vb2DblQ) & vbDblQ: End Function 'Quote @S as vbStr, which is quoting with double-quote and inside-double-quote will become 2 double-quote.
Function QuoBigBkt$(S): QuoBigBkt = "{" & S & "}": End Function
Function QuoBkt$(S): QuoBkt = "(" & S & ")": End Function
Function QuoDte$(S): QuoDte = QuoBy(S, "#"): End Function
Function QuoDot$(S): QuoDot = QuoBy(S, "."): End Function
Function QuoBig$(S):     QuoBig = vbBigOpn & S & vbBigCls: End Function
Function QuoDbl$(S):     QuoDbl = QuoBy(S, vbDblQ): End Function
Function QuoSng$(S):     QuoSng = QuoBy(S, "'"): End Function
Function QuoSpc$(S):     QuoSpc = QuoBy(S, " "): End Function
Function QuoSq$(S): QuoSq = "[" & S & "]": End Function
Function QuoSqAv(Av()) As String() 'Quote each element in @Av by square bracket and return it as string array.
Dim I: For Each I In Av
    PushI QuoSqAv, QuoSq(I)
Next
End Function

Function QuoBy$(S, By$): QuoBy = By & S & By: End Function

