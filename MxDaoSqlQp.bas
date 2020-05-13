Attribute VB_Name = "MxDaoSqlQp"
Option Compare Text
Option Explicit
Const CLib$ = "QSql."
Const CMod$ = CLib & "MxDaoSqlQp."
Const KwBet$ = "between"
Const KwSet$ = "set"
Const KwDis$ = "distinct"
Const KwUpd$ = "update"
Const KwInto$ = "into"
Const KwSel$ = "select"
Const KwFm$ = "from"
Const KwGp$ = "group by"
Const KwWh$ = "where"
Const KwAnd$ = "and"
Const KwOn$ = "on"
Const KwLJn$ = "left join"
Const KwIJn$ = "inner join"
Const KwOr$ = "or"
Const KwOrd$ = "order by"
Const KwLeftJn$ = "left join"
Const C_Sel$ = " " & KwSel & " "
Const C_Bet$ = " " & KwBet & " "
Const C_Ord$ = " " & KwOrd & " "
Const C_And$ = " " & KwAnd & " "
Const C_IJn$ = " " & KwIJn & " "
Const C_T$ = " "
Const C_Fm$ = " " & KwFm & " "
Const C_Set$ = " " & KwSet & " "
Const C_Dis$ = " " & KwDis & " "
Const C_Gp$ = " " & KwGp & " "
Const C_Wh$ = " " & KwWh & " "
Const C_Into$ = " " & KwInto & " "
Type FldMap: Extn As String: Intn As String: End Type ' Deriving(Ctor Ay)

Function JnAnd$(Ay): JnAnd = Jn(Ay, QuoSpc(KwAnd)): End Function
Function QpFm$(T, Optional Alias$): QpFm = C_Fm & QuoSq(T) & " " & Alias: End Function
Function QpFm_X$(X): QpFm_X = C_Fm & X: End Function
Function QpFmzX$(FmT_Using_AliX$): QpFmzX = QpFm(FmT_Using_AliX) & " x": End Function
Function QpGp$(Gp$): QpGp = AddPfxIfNB(Gp, C_Gp): End Function
Function QpGp_EprVblAy$(EprVblAy$()): QpGp_EprVblAy = C_Gp & JnCrLf(FmtEprVblAy(EprVblAy)): End Function
Function QpIns_T$(T): QpIns_T = "Insert into [" & T & "]": End Function
Function QpBkt_FF$(FF$): QpBkt_FF = QuoBkt(QpFis(FF)): End Function

Function QpBkt_Vy$(SqlVy)
Dim O$()
Dim V: For Each V In SqlVy
    PushI O, QuoSqlv(V)
Next
QpBkt_Vy = QuoBkt(JnComma(O))
End Function

Function QpInto_T$(T): QpInto_T = C_Into & "[" & T & "]": End Function
Function QpF12$(F12$): QpF12 = JnComma(SplitSpc(F12)): End Function
Function QpAnd$(Bexp$): QpAnd = AddPfxIfNB(Bexp, C_And): End Function
Function PFldInX_F_InAet_Wdt(F, Sset As Dictionary, Wdt%) As String()
Dim A$
    A = "[F] in ("
Dim I
'For Each I In LyJnQSqlCommaAetW(Sset, Wdt - Len(A))
    PushI PFldInX_F_InAet_Wdt, I
'Next
End Function

Function POnzJnXA(JnFny$())
Dim X$(): X = SyzQAy("x.[?]", JnFny)
Dim A$(): A = SyzQAy("a.[?]", JnFny)
Dim J$(): J = FmtAyab(X, A, " = ")
Dim S$: S = JnAnd(J)
POnzJnXA = KwOn & " " & S
End Function

Function QpOrd$(By$): QpOrd = AddPfxIfNB(By, C_Ord): End Function
Function QpOrd_DashSfxFF$(OrdMinusSfxFF$)
If OrdMinusSfxFF = "" Then Exit Function
Dim O$(): O = SyzSS(OrdMinusSfxFF)
Dim I, J%
For Each I In O
    If HasSfx(O(J), "-") Then
        O(J) = RmvSfx(O(J), "-") & " desc"
    End If
    J = J + 1
Next
QpOrd_DashSfxFF = QpOrd(JnCommaSpc(O))
End Function
Function QpDis$(IsDis As Boolean): QpDis = SzTrue(IsDis, C_Dis): End Function
Function QpSelStar$(): QpSelStar = KwSel & " *": End Function
Function QpSel_F$(F$, Optional IsDis As Boolean): QpSel_F = KwSel & QpDis(IsDis) & QuoSq(F): End Function
Function QpSel_FF$(FF$, Optional QpDis As Boolean): QpSel_FF = QpSel_Fny(FnyzFF(FF), QpDis): End Function
Function QpSel_FF_Extny$(FF$, Extny$()): QpSel_FF_Extny = QpSel_X(QpSel_Fny_Extny(Ny(FF), Extny)): End Function
Function QpSel_Fny$(Fny$(), Optional IsDis As Boolean): QpSel_Fny = KwSel & QpDis(IsDis) & JnCommaSpc(Fny): End Function
Function QpSel_Fny_Extny$(Fny$(), Extny$(), Optional IsDis As Boolean)
If Not ShdFmtSql Then QpSel_Fny_Extny = QpSel_Fny_Extny_NOFMT(Fny, Extny): Exit Function
Dim E$(), F$()
F = Fny
E = Extny
FEs_SetExtNm_ToBlnk_IfEqToFld F, E
FEs_SqQuoExtNm_IfNB E
FEs_AliExtNm E
FEs_AddAs_Or4Spc_ToExtNm E
FEs_AddTab2Spc_ToExtNm E
FEs_AliFld F
QpSel_Fny_Extny = KwSel & QpDis(IsDis) & C_NL & Join(FmtAyab(E, F), ",")
End Function

Function QpSel_Fny_Extny_NOFMT(Fny$(), Extny$(), Optional IsDis As Boolean)
Dim O$(), J%, E$, F$
For J = 0 To UB(Fny)
    F = Fny(J)
    E = Trim(Extny(J))
    Select Case True
    Case E = "", E = F: PushI O, F
    Case Else: PushI O, QuoSq(E) & " As " & F
    End Select
Next
QpSel_Fny_Extny_NOFMT = KwSel & QpDis(IsDis) & " " & JnCommaSpc(O)
End Function

Function QpSelStar_Fm$(T, Optional Bexp$): QpSelStar_Fm = C_Sel & "*" & QpFm(T) & Wh(Bexp): End Function

Function QpSelX_FldMapy$(M() As FldMap)
Dim O$()
Dim J%: For J = 0 To FldMapUB(M)
    PushI O, QpAs_FldMap(M(J))
Next
QpSelX_FldMapy = JnCommaSpc(O)
End Function

Function QpAs_FldMap$(M As FldMap)
With M
    QpAs_FldMap = QuoSq(.Extn) & " As " & .Intn
End With
End Function
Function QpSel_X$(X$, Optional IsDis As Boolean): QpSel_X = KwSel & QpDis(IsDis) & X: End Function

Function QpSet_FF_EprAy$(FF$, Ey$())
Const CSub$ = CMod & "QpSet_FF_EprAy"
Dim Fny$(): Fny = FnyzFF(FF)
Ass IsVblAy(Ey)
If Si(Fny) <> Si(Ey) Then Thw CSub, "[FF-Sz} <> [Si-Ey], where [FF],[Ey]", Si(Fny), Si(Ey), FF, Ey
Dim AFny$()
    AFny = AmAli(Fny)
    AFny = AmAddSfx(AFny, " = ")
Dim W%
    'W = VblWdty(Ey)
Dim Ident%
    W = WdtzAy(AFny)
Dim Ay$()
    Dim J%, U%, S$
    U = UB(AFny)
    For J = 0 To U
        If J = U Then
            S = ""
        Else
            S = ","
        End If
        'Push Ay, VblAli(Ey(J), Pfx:=AFny(J), IdentOpt:=Ident, WdtOpt:=W, Sfx:=S)
    Next
Dim Vbl$
    Dim Ay1$()
    Dim P$
    For J = 0 To U
        If J = 0 Then P = "|  Set" Else P = ""
'        Push Ay1, VblAli(Ay(J), Pfx:=P, IdentOpt:=6)
    Next
    Vbl = JnVBar(Ay1)
QpSet_FF_EprAy = Vbl
End Function

Function QpSet_FF_Ey$(FF$, Ey$()): QpSet_FF_Ey = QpSet_Fny_Ey(FnyzFF(FF), Ey): End Function


Function QpSet_Fny_Ey$(Fny$(), Ey$())
Dim J$(): J = FmtAyab(AmQuoSq(Fny), Ey, " = ")
Dim J1$(): J1 = AmAddPfx(J, C_T)
Dim S$: S = Jn(J, "," & C_NL)
QpSet_Fny_Ey = C_NLT & KwSet & C_NL & S
End Function

Function QpSet_Fny_Vy$(Fny$(), Vy())
Dim F$(): F = AmQuoSq(Fny)
Dim V$(): V = QuoSqlvy(Vy)
QpSet_Fny_Vy = JnComma(FmtAyab(F, V, "="))
End Function

Function QpSet_Fny_Vy1$(Fny$(), Vy())
Dim A$: GoSub X_A
QpSet_Fny_Vy1 = "  Set " & A
Exit Function
X_A:
    Dim L$(): L = AmQuoSq(Fny)
    Dim R$(): R = QuoSqlvy(Vy)
    Dim J%, O$()
    For J = 0 To UB(L)
        Push O, L(J) & " = " & R(J)
    Next
    A = JnComma(O)
    Return
End Function

Function QpSetzX$(SetX$): QpSetzX = C_Set & SetX: End Function

Function QpSetzXA(FnyX$(), FnyA$())
Dim X$(): X = AmAddPfxSfx(FnyX, "x.[", "]")
Dim A$(): A = AmAddPfxSfx(FnyA, "a.[", "]")
Dim J$(): J = FmtAyab(X, A, " = ")
          J = AmAddPfx(J, C_T)
Dim S$:   S = Jn(J, "," & C_NL)
QpSetzXA = QpSetzX(S)
End Function

Function QpSetzXAFny(Fny$()): QpSetzXAFny = QpSetzXA(Fny, Fny): End Function
Function QpIJn_T_A_Fny(TblX$, TblA$, JnFny$()): QpIJn_T_A_Fny = C_T & "[" & TblX & "] x" & C_IJn & "[" & TblA & "] a " & POnzJnXA(JnFny): End Function
Function QpUpd$(T): QpUpd = KwUpd & C_T & QuoT(T): End Function
Function QpUpd_X$(TblX$): QpUpd_X = KwUpd & C_NL & TblX: End Function
Function QpUpd_X_A_Jn$(TblX$, TblA$, JnFny$()): QpUpd_X_A_Jn = QpUpd_X(QpIJn_T_A_Fny(TblX, TblA, JnFny)): End Function

Function WhFeq(F$, Eqval, Optional Alias$): WhFeq = Wh(QuoF(F, Alias) & "=" & QuoSqlv(Eqval)): End Function
Function Wh_F_In$(F$, InVy): Wh_F_In = Wh(Bexp_F_In(F, InVy)): End Function
Function Wh_FF_Eq$(FF$, EqVy): Wh_FF_Eq = Wh_Fny_Eq(FnyzFF(FF), EqVy): End Function
Function Wh_Fny_Eq$(Fny$(), EqVy): Wh_Fny_Eq = Wh(Bexp_Fny_EqVy(Fny, EqVy)): End Function
Function Wh_T_EqK$(T, K&): Wh_T_EqK = WhFeq(T & "Id", K): End Function
Function Wh_T_Id$(T, Id): Wh_T_Id = Wh(FmtQQ("[?]Id=?", T, Id)): End Function
Function WhBet_F_Fm_To$(F$, FmV, ToV, Optional Alias$): WhBet_F_Fm_To = C_Wh & QuoF(F, Alias) & C_Bet & QuoT(FmV) & C_And & QuoT(ToV): End Function

Private Sub QpGp_EprVblAy__Tst()
Dim EprVblAy$()
    Push EprVblAy, "1lskdf|sdlkfjsdfkl sldkjf sldkfj|lskdjf|lskdjfdf"
    Push EprVblAy, "2dfkl sldkjf sldkdjf|lskdjfdf"
    Push EprVblAy, "3sldkfjsdf"
DmpAy SplitVBar(QpGp_EprVblAy(EprVblAy))
End Sub

Private Sub QpSel__Tst()
Dim Fny$(), EprVblAy$()
EprVblAy = Sy("F1-Epr", "F2-Epr   AA|BB    X|DD       Y", "F3-Epr  x")
Fny = SplitSpc("F1 F2 F3xxxxx")
'Debug.Print LineszVbl(QpSelFFFldLvs(Fny, EprVblAy))
End Sub

Private Sub QpSel_Fny_Extny__Tst()
Dim Fny$()
Dim Extny$()
GoSub Z
Exit Sub
Z:
    Fny = SyzSS("Sku CurRateAc VdtFm VdtTo HKD Per CA_Uom")
    Extny = Termy("Sku [     Amount] [Valid From] [Valid to] Unit per Uom")
    Debug.Print QpSel_Fny_Extny(Fny, Extny)
    Return
End Sub

Private Sub QpSet_Fny_VyFmt__Tst()
Dim Fny$(), Vy()
Ept = LineszVbl("|  Set|" & _
"    [A xx] = 1                     ,|" & _
"    B      = '2'                   ,|" & _
"    C      = #2018-12-01 12:34:56# ")
Fny = Termy("[A xx] B C"): Vy = Array(1, "2", #12/1/2018 12:34:56 PM#): GoSub Tst
Exit Sub
Tst:
    Act = QpSet_Fny_Vy(Fny, Vy)
    C
    Return
End Sub

Private Sub QpSetFFEqvy__Tst()
Dim Fny$(), EprVblAy$()
Fny = SyzSS("a b c d")
Push EprVblAy, "1sdfkl|lskdfj|skldfjskldfjs dflkjsdf| sdf"
Push EprVblAy, "2sdfkl|lskdfjdf| sdf"
Push EprVblAy, "3sdfkl|fjskldfjs dflkjsdf| sdf"
Push EprVblAy, "4sf| sdf"
    'Act = QpSet_Fny_Evy(Fny, EprVblAy)
'Debug.Print LineszVbl(Act)
End Sub

Private Sub Wh_F_In__Tst()
Dim F$, Vy()
F = "A"
Vy = Array(1, "2", #2/1/2017#)
Ept = " where A=1 and B='2' and C=#2017-2-1#"
GoSub Tst
Exit Sub
Tst:
    Act = Wh_F_In(F, Vy)
    C
    Return
End Sub

Property Get C_NL$() ' New Line
If ShdFmtSql Then
    C_NL = vbCrLf
Else
    C_NL = " "
End If
End Property

Property Get C_NLT$() ' New Line Tabe
If ShdFmtSql Then
    C_NLT = C_NL & C_T
Else
    C_NLT = " "
End If
End Property

Property Get C_NLTT$() ' New Line Tabe
If ShdFmtSql Then
    C_NLTT = C_NLT & C_T
Else
    C_NLTT = " "
End If
End Property


Function QpSelStarInto$(Into): QpSelStarInto = "Select * Into [" & Into & "]": End Function
Function QpInto$(Into$)
QpInto = vbCrLf & " Into [" & Into & "]"
End Function
Function QpIntoFm$(Into$, Fm$)
QpIntoFm = vbCrLf & " Into [" & Into & "] From [" & Fm & "]"
End Function

'-----------------------------------------------------------------------------------
Function Wh$(Bexp$): Wh = AddPfxIfNB(Bexp, C_Wh): End Function
Function QpGp_FF$(FF$): QpGp_FF = vbCrLf & " Group By " & QpFis(FF): End Function
Function QpSelDist$(F$, T$): QpSelDist = "Select Distinct [" & F & "] From [" & T & "]": End Function
Function QpAndFeq$(F$, Eqval, Optional Alias$): QpAndFeq = vbCrLf & " and " & Bexp_F_Eq(F, Eqval, Alias): End Function
Function QpInsInto$(T): QpInsInto = "Insert Into [" & T & "]": End Function
'---========================================================================== Qp*
Function QpFldLis$(FF$): QpFldLis = JnComma(AmQuoSq(FnyzFF(FF))): End Function
Function QpValues$(Dr): QpValues = JnComma(QuoSqlvy(Dr)): End Function
Function QpFis$(FF$, Optional Alias$)
':Fis: :Csvln ! #Fld-List# F=Fld; is=List;
':Csvln: :Ln  ! #Comma-Separated-Value-Line#
Dim A$(): A = QuoTermy(Termy(FF))
If Alias <> "" Then A = AmAddPfx(A, Alias & ".")
QpFis = JnCommaSpc(A)
End Function
Function QpBfis$(FF$): QpBfis = QuoBkt(QpFis(FF)): End Function
Function QpIJn(T$, FF$, Optional AliasX$ = "x", Optional AliasA$ = "a"): QpIJn = QpJn(KwIJn, T, FF, AliasX, AliasA): End Function
Function QpLJn(T, FF$, Optional AliasX$ = "x", Optional AliasA$ = "a"): QpLJn = QpJn(KwLJn, T, FF, AliasX, AliasA): End Function

Private Function QpJn(KwJn$, T, FF$, AliasX$, AliasA$)
Dim X$, A$, O$()
Dim F: For Each F In FnyzFF(FF)
    X = QuoT(F, AliasX)
    A = QuoT(F, AliasA)
    PushI O, FmtQQ("? = ?", A, X)
Next
Dim TT$: TT = QuoSpc(QuoSq(T)) & AliasA & " "
QpJn = C_NLT & KwJn & TT & JnAnd(O) & ")"
End Function
Function FldMap(Extn, Intn) As FldMap
With FldMap
    .Extn = Extn
    .Intn = Intn
End With
End Function
Function AddFldMap(A As FldMap, B As FldMap) As FldMap(): PushFldMap AddFldMap, A: PushFldMap AddFldMap, B: End Function
Sub PushFldMapAy(O() As FldMap, A() As FldMap): Dim J&: For J = 0 To FldMapUB(A): PushFldMap O, A(J): Next: End Sub
Sub PushFldMap(O() As FldMap, M As FldMap): Dim N&: N = FldMapSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function FldMapSi&(A() As FldMap): On Error Resume Next: FldMapSi = UBound(A) + 1: End Function
Function FldMapUB&(A() As FldMap): FldMapUB = FldMapSi(A) - 1: End Function
