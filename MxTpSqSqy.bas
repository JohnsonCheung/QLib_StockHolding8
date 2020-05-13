Attribute VB_Name = "MxTpSqSqy"
Option Explicit
Option Compare Text
#If False Then
Const CLib$ = "QTp."
Const CMod$ = CLib & "MxTpSqy."
Enum eSqBlkTy: eErBlk: eSqBlk: ePmBlk: eSwBlk: End Enum
Enum eSqSpect: eSelStmt: eUpdStmt: eDrpStmt: End Enum
Public Const eSqSpect$ = "SelStmt UpdStmt DrpStmt"
Const KwInto$ = "INTO"
Const KwSel$ = "SEL"
Const KwSelDis$ = "SELECT DISTINCT"
Const KwFm$ = "FM"
Const KwGp$ = "GP"
Const KwWh$ = "WH"
Const KwAnd$ = "AND"
Const KwJn$ = "JN"
Const KwLeftJn$ = "LEFT JOIN"
Const SqTpKwSS$ = "Pm Sw Sql Rmk"
Type PmSw2
    Pm As Dictionary
    StmtSw As Dictionary
    FldSw As Dictionary
End Type

Function SqyzD(B As SqBld) As String()

End Function
Private Function SqlnErn$(A As LLn, B As XPmSw2)
Dim S As S12: S = BrkSpc(A.Ln)
Select Case True
Case HasPfx(S.S1, ">?")
    Select Case S.S2
    Case "0", "1"
    Case SqlnErn = LLnMsg(A, "If K='>?xxx', V should be 0 or 1.")
    End Select
Case HasPfx(S.S1, ">")
Case Else: SqlnErn = LLnMsg(A, FmtQQ("Pm line should beg with (>? | >)"))
End Select
End Function

Private Function LLnMsg$(A As LLn, Msg$)

End Function

Private Sub SqyzTpLy__Tst()
Dim SqTpLy$()
GoSub Z
Exit Sub
Z:
    B SqyzTpLy(SampSqTpLy)
    Return
End Sub

Function SqyzTpLy(SqTpLy$()) As String()
Const CSub$ = "SqyFmTp"
Dim S As SqTpSrc: S = SqTpSrczT(SqTpLy)
ChkEr SqTpEr(S), CSub
SqyzTpLy = UUSqy(UUBld(UUDta(S)))
End Function

Function SqlSel$(Sel$(), EDic As Dictionary, FldSw As Dictionary)
'BrwKLys Sel
Dim LFm$, LInto$, LSel$, LOrd$, LWh$, LGp$, LAndOr$(), LAlias$()
'    LSel = ShfKLyMLin(X, "Sel")
'    LInto = ShfKLyMLin(X, "Into")
'    LFm = ShfKLyMLin(X, "Fm")
'    LJn = ShfKLyMLyzKK(X, "Jn LJn")
'    LWh = ShfKLyOLin(X, "Wh")
'    LAndOr = ShfKLyMLyzKK(X, "And Or")
'    LGp = ShfKLyOLin(X, "Gp")
'    LOrd = ShfKLyOLin(X, "Ord")
Dim ADic As Dictionary: Set ADic = DiczVkkLy(LAlias)
Dim Ffny$(), FGp$()
'    Ffny = SQ_SelFny(LSel, FldSw)
    FGp = QpSelFld(LGp, FldSw)
Dim OX$, Into$, OT$, OWh$, OGp$, OOrd$
'    Dim Fny$()

'    '
    Into = RmvT1(LInto)
Stop '    OGp = QpEprLis(FGp, EDic, ADic)
'    OOrd = SQ_EprLis(FOrd, EDic, ADic)
'    OWh = SQ_Wh()
'    OT = RmvT1(LFm)
'SqlSel = SqlSel_X_Into_T_Wh_Gp_Ord(OX, Into, OT, OGp, OOrd)
End Function

Function QpSelFld(FF$, FldSw As Dictionary) As String()
Const CSub$ = CMod & "SQ_SelFld"
Dim Fny$(): Fny = FnyzFF(FF)
Dim F1$, F
For Each F In Fny
    F1 = FstChr(F)
    Select Case True
    Case F1 = "?"
        If Not FldSw.Exists(F) Then Thw CSub, "An option fld not found in FldSw", "Opt-Fld FF FldSw", F, FF, FldSw
        If FldSw(F) Then
            'PushI XFny, RmvFstChr(F)
        End If
    Case F1 = "$"
        'PushI XFny, RmvFstChr(F)
    Case Else
        'PushI XFny, F
    End Select
Next
Stop
End Function

Private Function IsSkip(FstSqln$, Sqlny$(), T As eSqSpect, StmtSw As Dictionary) As Boolean
Const CSub$ = CMod & "IsSkip"
If FstChr(FstSqln) <> "?" Then Exit Function
Dim Key$: Key = StmtSwk(Sqlny, T)
If Not StmtSw.Exists(Key) Then Thw CSub, "StmtSw does not contain the StmtSwk", "Sqlny StmtSwk StmtSw", Sqlny, Key, StmtSw
IsSkip = Not StmtSw(Key)
End Function

Function SqyzL(L() As SqDtaln, P As PmSw2) As String()
Dim J%: For J = 0 To SqSrclnUB(L)
    PushI SqyzL, SqlzL(L(J), P)
Next
End Function

Function SqlzSp$(Sqsp, P As PmSw2)
Dim Sqlny$(): Sqlny = Sqsp
Dim FstSqln$:    FstSqln = Sqlny(0)
Dim Ty As eSqSpect:     Ty = SqSpect(FstSqln)
Dim Skip As Boolean:  Skip = IsSkip(FstSqln, Sqlny, Ty, P.Sw2.StmtSw)
                             If Skip Then Exit Function
Dim S$():                S = SQ_RmvEprLin(Sqlny)
Dim E As Dictionary: Set E = SQ_EprDic(Sqlny)
Dim O$
    Select Case True
    Case Ty = eDrpStmt: O = sqlDrp(S)
    Case Ty = eUpdStmt: O = SqlUpd(S, E, P.Sw2.FldSw)
    Case Ty = eSelStmt: O = SqlSel(S, E, P.Sw2.FldSw)
    Case Else: Imposs CSub
    End Select
SqlzSp = O
End Function

Private Function sqlDrp$(Drp$())
End Function
Private Function SqlUpd$(Upd$(), EDic As Dictionary, FldSw As Dictionary)
End Function



Function SQ_RmvEprLin(Sqlny$()) As String()
SQ_RmvEprLin = AePfx(Sqlny, "$")
End Function


Function SQ_EprDic(Sqlny$()) As Dictionary
Set SQ_EprDic = Dic(CvSy(AwPfx(Sqlny, "$")))
End Function

Function SqSpect(FstSqln$) As eSqSpect
Dim L$: L = RmvPfx(T1(FstSqln), "?")
Select Case L
Case "SEL", "SELDIS": SqSpect = eSelStmt
Case "UPD": SqSpect = eUpdStmt
Case "DRP": SqSpect = eDrpStmt
Case Else: Stop
End Select
End Function

Function StmtSwk$(Sqlny$(), T As eSqSpect) ' #statment-switch-key#
Dim O$
Select Case T
Case eSelStmt: O = StmtSwkzSel(Sqlny)
Case eUpdStmt: O = StmtSwkzUpd(Sqlny)
Case Else:  EnmEr CSub, "eSqSpect", eSqSpectSS, T
End Select
StmtSwk = "?:" & O
End Function

Function StmtSwkzSel$(SelSqlny$())
StmtSwkzSel = FstElewRmvT1(SelSqlny, "into")
End Function

Function StmtSwkzUpd$(UpdSqlny$())
Dim Lin1$
    Lin1 = UpdSqlny(0)
If RmvPfx(ShfTerm(Lin1), "?") <> "upd" Then Stop
StmtSwkzUpd = Lin1
End Function

Function IsXXX(A$(), XXX$) As Boolean
IsXXX = UCase(T1(A(UB(A)))) = XXX
End Function

Function SQ_And(A$(), E As Dictionary)
'and f bet xx xx
'and f in xx
Dim F$, I, L$, Ix%
For Each I In Itr(A)
    'Set M = I
    'LnxAsg M, L, Ix
    If ShfTerm(L) <> "and" Then Stop
    F = ShfTerm(L)
    Select Case ShfTerm(L)
    Case "bet":
    Case "in"
    Case Else: Stop
    End Select
Next
End Function

Function SQ_Gp$(GG$, FldSw As Dictionary, E As Dictionary)
If GG = "" Then Exit Function
Dim EprAy$(), Ay$()
Stop
'    EprAy = DicSelIntoSy(EDic, Ay)
'XGp = SqpGp(EprAy)
End Function

Function SQ_JnOrLeftJn(A$(), E As Dictionary) As String()

End Function

Function SQ_Sel$(A$, E As Dictionary)
Dim Fny$()
    Dim T1$, L$
    L = A
    T1 = RmvPfx(ShfTerm(L), "?")
    'Fny = XSelFny(SyzSS(L), FldSw)
Select Case T1
'Case KwSel:    XSel = X.Sel_FnSampEDic(Fny, E)
'Case KwSelDis: XSel = X.Sel_FnSampEDic(Fny, E, IsDis:=True)
Case Else: Stop
End Select
End Function
Function SQ_SelFny(Fny$(), FldSw As Dictionary) As String()
Dim F
For Each F In Fny
    If FstChr(F) = "?" Then
        If Not FldSw.Exists(F) Then Stop
        'If FldSw(F) Then PushI XSelFny, F
    Else
        'PushI XSelFny, F
    End If
Next
End Function

Function SQ_Set(DroLLn(), E As Dictionary, OEr$())

End Function

Function SQ_Upd(DroLLn(), E As Dictionary, OEr$())

End Function
Function SQ_Wh$() ' (L$, E As Dictionary)
'L is following
'  ?Fld in @ValLis  -
'  ?Fld bet @V1 @V2
Dim F$, Vy$(), V1, V2, IsBet As Boolean
If IsBet Then
'    If Not FndValPair(F, E, V1, V2) Then Exit Function
    'XWh = SWhBet(F, V1, V2)
    Exit Function
End If
'If Not FndVy(F, E, Vy, Q) Then Exit Function
'XWh = SWhFldInVSampStr(F, Vy)
End Function

Function SQ_WhBetNum$(DroLLn(), E As Dictionary, OEr$())

End Function

Function SQ_WhEpr(DroLLn(), E As Dictionary, OEr$())

End Function

Function SQ_WhInNumLis$(DroLLn(), E As Dictionary, OEr$())

End Function

Function CvVSampToTF_Fm01(A As Dictionary) As Dictionary
Dim O As Dictionary: Set O = CloneDic(A)
Dim K
For Each K In O.Keys
    Select Case O(K)
    Case "0": O(K) = False
    Case "1": O(K) = True
    End Select
Next
Set CvVSampToTF_Fm01 = O
End Function

Private Sub sqlSel__Tst()
Dim E As Dictionary, Ly$(), FldSw As Dictionary

'---
Erase XX
    X "?XX Fld-XX"
    X "BB Fld-BB-LINE-1"
    X "BB Fld-BB-LINE-2"
    Set E = Dic(XX)           '<== Set EprDic
Erase XX
    X "?XX 0"
    Set FldSw = Dic(XX)
    Set FldSw = CvVSampToTF_Fm01(FldSw)
Erase XX
    X "sel ?XX BB CC"
    X "into #AA"
    X "fm   #AA"
    X "jn   #AA"
    X "jn   #AA"
    X "wh   A bet $a $b"
    X "and  B in $c"
    X "gp   D C"        '<== LySq
GoSub Tst
Exit Sub
Tst:
    Act = SqlSel(Ly, E, FldSw)
    C
    Return
End Sub

Private Sub EprDic__Tst()
Dim Ly$()
Dim D As New Dictionary
'-----

Erase Ly
PushI Ly, "aaa bbb"
PushI Ly, "111 222"
PushI Ly, "$"
PushI Ly, "A B0"
PushI Ly, "A B1"
PushI Ly, "A B2"
PushI Ly, "B B0"
D.RemoveAll
    D.Add "A", JnCrLf(SyzSS("B0 B1 B2"))
    D.Add "B", "B0"
    Set Ept = D
GoSub Tst
Exit Sub
Tst:
    Set Act = SQ_EprDic(Ly)
    Ass IsEqDic(CvDic(Act), CvDic(Ept))
    
    Return
End Sub

Private Sub StmtSwk__Tst()
Dim Ly$(), Ty As eSqSpect
GoSub T0
GoSub T1
Exit Sub
'---
T0:
    Erase Ly
    PushI Ly, "sel sdflk"
    PushI Ly, "fm AA BB"
    PushI Ly, "into XX"
    Ept = "XX"
    Ty = eSelStmt
    GoTo Tst
T1:
    Erase Ly
    PushI Ly, "?upd XX BB"
    PushI Ly, "fm dsklf dsfl"
    Ept = "XX BB"
    Ty = eUpdStmt
    GoTo Tst
Tst:
    Act = StmtSwk(Ly, Ty)
    C
    Return
End Sub
#End If
