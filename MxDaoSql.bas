Attribute VB_Name = "MxDaoSql"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoSql."

Sub FEs_AddAs_Or4Spc_ToExtNm(OE$())
Dim J%, C$
For J = 0 To UB(OE)
    
    If Trim(OE(J)) = "" Then
        C = "    "
    Else
        C = " As "
    End If
    OE(J) = OE(J) & C
Next
End Sub

Sub FEs_AddTab2Spc_ToExtNm(OE$()): OE = AmAddPfx(OE, "  "): End Sub
Sub FEs_AliExtNm(OE$()): OE = AmAli(OE): End Sub
Sub FEs_AliFld(OF$()): OF = AmAli(OF): End Sub
Sub FEs_SetExtNm_ToBlnk_IfEqToFld(F$(), OE$())
Dim J%: For J = 0 To UB(OE)
    If OE(J) = F(J) Then OE(J) = ""
Next
End Sub

Sub FEs_SqQuoExtNm_IfNB(OE$())
Dim J%: For J = 0 To UB(OE)
    If OE(J) <> "" Then OE(J) = QuoSq(OE(J))
Next
End Sub

Function SqlAddCol_T_Fny_FzDiSqlTy$(T, Fny$(), FzDiSqlTy As Dictionary)
Dim O$()
Dim F: For Each F In Fny
    PushI O, F & " " & VzDicK(FzDiSqlTy, F, "FzDiSqlTy", "Fld")
Next
SqlAddCol_T_Fny_FzDiSqlTy = FmtQQ("Alter Table [?] add column ?", T, JnComma(O))
End Function

Function SqlAddColzAy$(T, ColAy$()): SqlAddColzAy = SqlAddColzLis(T, JnCommaSpc(ColAy)): End Function
Function SqlAddColzLis$(T, ColLis$): SqlAddColzLis = FmtQQ("Alter Table [?] add column ?", T, ColLis): End Function
Function SqlCrtTbl_T_X$(T, X$): SqlCrtTbl_T_X = FmtQQ("Create Table [?] (?)", T, X): End Function
Function SqlDlt$(T, Optional Bexp$): SqlDlt = SqlDlt_T(T, Bexp): End Function
Function SqlDlt_T$(T, Optional Bexp$): SqlDlt_T = "Delete * from [" & T & "]" & Wh(Bexp): End Function
Function SqlDrpCol$(T, FF$): SqlDrpCol = FmtQQ("Alter Table [?] drop column ?", T, QpFis(FF)): End Function
Function SqlDrpCol_T_F$(T, F$): SqlDrpCol_T_F = FmtQQ("Alter Table [?] drop column [?]", T, F$): End Function
Function SqlDrpFld$(T, Fny$()): SqlDrpFld = "Alter Table [" & T & "] drop column " & JnCommaSpc(AmQuoSq(Fny)): End Function
Function SqlDrpTbl_T$(T): SqlDrpTbl_T = "Drop Table [" & T & "]": End Function
Function SqlIns_T_FF_Dr$(T, FF$, Dr): SqlIns_T_FF_Dr = FmtQQ("Insert Into [?] (?) Values(?)", T, QpFldLis(FF), QpValues(Dr)): End Function
Function SqlIns_T_FF_Vap$(T, FF$, ParamArray Vap())
Dim Vy(): Vy = Vap
SqlIns_T_FF_Vap = QpIns_T(T) & QpBkt_FF(FF) & " Values" & QpBkt_Vy(Vy)
End Function

Function SqlSel_F$(F$): SqlSel_F = SqlSel_F_T(F, F): End Function
Function SqlSel_F12_Fm(F12$, T, Optional Bexp$): SqlSel_F12_Fm = SqlSel_X_Fm(QpF12(F12), T, Bexp): End Function
Function SqlSel_F_T_WhF_In(F, T, WhF, InVy): SqlSel_F_T_WhF_In = SqlSel_F_T(F, T, Bexp_F_In(WhF, InVy)): End Function
Function SqlSel_F_T_WhF_Eq(F, T, WhF, Eqval): SqlSel_F_T_WhF_Eq = SqlSel_F_T(F, T, Bexp_F_Eq(WhF, Eqval)): End Function
Function SqlSel_F_T_WhFny_Eq(F, T, WhFny$(), EqVy): SqlSel_F_T_WhFny_Eq = SqlSel_F_T(F, T, Bexp_Fny_EqVy(WhFny, EqVy)): End Function
Function SqlSel_F_T$(F, T, Optional Bexp$, Optional Dis As Boolean): SqlSel_F_T = FmtQQ("Select ?[?] from [?]?", QpDis(Dis), F, T, Wh(Bexp)): End Function
Function SqlSel_TF$(TF$, Optional Bexp$, Optional Dis As Boolean)
Dim A As S12: A = BrkTF(TF)
SqlSel_TF = SqlSel_F_T(A.S2, A.S1, Bexp, Dis)
End Function

Function SqlSel_FF_EDic_Into_T$(FF$, EDic As Dictionary, Into, T, Optional Bexp$)
Dim Fny$(): Fny = FnyzFF(FF)
Dim EprAy$(): EprAy = SyzDicKy(EDic, Fny)
SqlSel_FF_EDic_Into_T = SqlSel_Fny_Extny_Into_Fm(Fny, EprAy, Into, T, Bexp)
End Function

'Function SqlSel_FF_EprDic_T$(FF$, E As Dictionary, T, Optional Dis As Boolean): SqlSel_FF_EprDic_T = "Select" & vbCrLf & FFEprDicAsLines(FF$, E): End Function
Function SqlSel_FF_X_Wh$(FF$, X$, Bexpr$): SqlSel_FF_X_Wh = QpSel_FF(FF) & QpFm_X(X) & Wh(Bexpr): End Function
Function SqlSel_FF_T$(FF$, T, Optional Dis As Boolean, Optional Bexp$): SqlSel_FF_T = QpSel_FF(FF, Dis) & QpFm(T) & Wh(Bexp): End Function
Function SqlSel_FF_T_Ord(FF$, T, OrdMinusSfxFF$):                       SqlSel_FF_T_Ord = QpSel_FF(FF) & QpFm(T) & QpOrd_DashSfxFF(OrdMinusSfxFF): End Function
Function SqlSel_FF_T_Ordff$(FF$, T, OrdMinusSfxFF$):                  SqlSel_FF_T_Ordff = QpSel_FF(FF) & QpFm(T) & QpOrd_DashSfxFF(OrdMinusSfxFF): End Function
Function SqlSel_FF_T_WhF_InVy$(FF$, T, WhF$, InVy, Optional Dis As Boolean): SqlSel_FF_T_WhF_InVy = SqlSel_FF_T(FF, T, Dis, Bexp_F_In(WhF$, InVy)): End Function
Function SqlSel_Fny_Extny_Into_Fm$(Fny$(), Extny$(), Into, Fm, Optional Bexp$): SqlSel_Fny_Extny_Into_Fm = QpSel_Fny_Extny(Fny, Extny) & QpInto_T(Into) & QpFm(Fm) & Wh(Bexp): End Function
Function SqlSel_Fny_T(Fny$(), T, Optional Bexp$, Optional Dis As Boolean): SqlSel_Fny_T = QpSel_Fny(Fny, Dis) & QpFm(T) & Wh(Bexp): End Function
Function SqlSel_Fny_Fm_WhFny_Eq$(Fny$(), T, WhFny$(), EqVy): SqlSel_Fny_Fm_WhFny_Eq = SqlSel_Fny_T(Fny, T, Wh_Fny_Eq(WhFny, EqVy)): End Function
Function SqlSelStar_Into_Fm_WhFalse$(Into, Fm): SqlSelStar_Into_Fm_WhFalse = FmtQQ("Select * Into [?] from [?] where false", Into, Fm): End Function
Function SqlSelStar_Fm$(T, Optional Bexp$):                             SqlSelStar_Fm = QpSelStar_Fm(T, Bexp):              End Function
Function SqlSelStar_Fm_F_Eq$(T, F, Eqval):                         SqlSelStar_Fm_F_Eq = QpSelStar_Fm(T, Bexp_F_Eq(F, Eqval)):    End Function
Function SqlSelStar_Fm_Fny_EqVy$(T, Fny$(), EqVy):             SqlSelStar_Fm_Fny_EqVy = QpSelStar_Fm(T, Bexp_Fny_EqVy(Fny, EqVy)):    End Function
Function SqlSelStar_T_Skvy$(D As Database, T, Skvy()):              SqlSelStar_T_Skvy = SqlSelStar_Fm_Fny_EqVy(T, SkFny(D, T), Skvy): End Function
Function SqlSel_F_Fm$(F$, Fm, Optional Bexp$, Optional Dis As Boolean):   SqlSel_F_Fm = QpSel_F(F, Dis) & QpFm(Fm) & Wh(Bexp): End Function
Function SqlSel_T_WhId$(T, Id&):                                        SqlSel_T_WhId = QpSelStar_Fm(T) & Wh_T_Id(T, Id):        End Function
Function SqlSel_X_Into_Fm$(X$, Into$, Fm$, Optional Bexp$, Optional Gp$, Optional Ord$, Optional Dis As Boolean): SqlSel_X_Into_Fm = QpSel_X(X) & QpInto_T(Into) & QpFm(Fm) & Wh(Bexp) & QpGp(Gp) & QpOrd(Ord$): End Function
Function SqlSel_X_Fm$(X$, T, Optional Bexp$): SqlSel_X_Fm = QpSel_X(X) & QpFm(T) & Wh(Bexp):                End Function
Function SqlSelCnt_Fm$(Fm, Optional Bexp$):  SqlSelCnt_Fm = "select Count(*) from [" & Fm & "]" & Wh(Bexp): End Function
Function SqlSel_FF_Fm$(FF$, Fm, Optional Dis As Boolean, Optional Bexp$): SqlSel_FF_Fm = SqlSel_FF_T(FF$, Fm, Dis:=True) & Wh(Bexp): End Function
Function SqlSelStar_Into_Fm$(Into$, Fm$, Optional Bexp$): SqlSelStar_Into_Fm = QpSelStar & QpInto_T(Into) & QpFm(Fm) & Wh(Bexp): End Function
Function SqlSel_Into_FF_Fm$(Into$, FF$, Fm$, Optional Bexp$, Optional Dis As Boolean): SqlSel_Into_FF_Fm = QpSel_FF(FF, Dis) & QpInto_T(Into) & QpFm(Fm) & Wh(Bexp): End Function
'Function SqlUpd_T_FF_EqDr_Whff_Eqvy$(T, FF$, Dr, WhFF$, EqVy): SqlUpd_T_FF_EqDr_Whff_Eqvy = QpUpd(T) & QpSet_FF_Eq(FF, Dr) & Wh_FF_Eq(WhFF, EqVy): End Function

Function SqlUpd_T_Sk_Fny_Dr$(T, Sk$(), Fny$(), Dr)
If Si(Sk) = 0 Then Stop
Dim QpUpd$, Set_$, Wh$: GoSub X_QpUpd_Set_Wh
'UpdSql = QpUpd & Set_ & Wh
Exit Function
X_QpUpd_Set_Wh:
    Dim Fny1$(), Dr1(), Skvy(): GoSub X_Fny1_Dr1_SkVy
    QpUpd = "Update [" & T & "]"
    Set_ = QpSet_Fny_Vy(Fny1, Dr1)
    Wh = Wh_Fny_Eq(Sk, Skvy)
    Return
X_Ay:
    Dim L$(), R$()
    L = AmAliQuoSq(Fny)
    R = QuoSqlvy(Dr)
    Return
X_Fny1_Dr1_SkVy:
    Dim Ski, J%, Ixy%(), I%
    For Each Ski In Sk
'        I = IxzAy(Fny, Ski)
        If I = -1 Then Stop
        Push Ixy, I
        Push Skvy, Dr(I)    '<====
    Next
    Dim F
    For Each F In Fny
        If Not HasEle(Ixy, J) Then
            Push Fny1, F        '<===
            Push Dr1, Dr(J)     '<===
        End If
        J = J + 1
    Next
    Return
End Function

Function SqlUpd_T_Fny_Ey$(T, Fny$(), Ey$(), Optional Bexp$): SqlUpd_T_Fny_Ey = QpUpd(T) & QpSet_Fny_Ey(Fny, Ey) & Wh(Bexp): End Function

Function SqlUpd_T_Fm_Jn_Set$(T$, FmA$, JnFny$(), SetFny$())
'Fm T     : Table nm to be update.  It will have alias x.
'Fm FmA   : Table nm used to update @T.  It will has alias a.
'Fm JnFny : Fld nm common in @T & @FmA.  It will use to bld the jn clause with alias x and a.
'Fm SetX  : Fny in @T to be updated.  No alias, by the ret sql will put the alias x.  Sam ele as @EqA.
'Ret      : upd sql stmt updating @T from @FmA using @JnFny as jn clause setting @T fld as stated in @SetX eq to @FmA fld as stated in @EqA
Dim U$: U = QpUpd_X_A_Jn(T, FmA, JnFny)
Dim S$: S = QpSetzXAFny(SetFny)
SqlUpd_T_Fm_Jn_Set = U & C_NL & S
End Function

Function SqlUpd_X_SetX$(TbX$, SetX$): SqlUpd_X_SetX = QpUpd_X(TbX) & QpSetzX(SetX): End Function

Function SqpAEqB_Fny_AliasAB$(Fny$(), Optional AliasAB$ = "x a")
Dim A1$: A1 = BefSpc(AliasAB) ' Alias1
Dim A2$: A2 = BefSpc(AliasAB) ' Alias2
Dim A$(): A = AmAddPfx(Fny, A1 & ".")
Dim B$(): B = AmAddPfx(Fny, A2 & ".")
Dim J$(): J = FmtAyab(A, B, " = ")
SqpAEqB_Fny_AliasAB = JnCommaSpc(J)
End Function

Function SqyCrtPkzTny(Tny$()) As String()
Dim T: For Each T In Itr(Tny)
    PushI SqyCrtPkzTny, sqlCrtPk(T)
Next
End Function

Function SqyDlt_T_WhFld_InAet(T, F, Sset As Dictionary, Optional SqlWdt% = 3000) As String()
Dim A$
Dim Ey$()
    A = SqlDlt_T(T) & " Where "
    Ey = PFldInX_F_InAet_Wdt(F, Sset, SqlWdt - Len(A))
Dim E
For Each E In Ey
    PushI SqyDlt_T_WhFld_InAet, A & E & vbCrLf
Next
End Function

Private Sub SqlSel_Fny_Ey_Into_T_OB__Tst()
Dim Fny$(), Ey$(), Into$, T$, Bexp$
GoSub Z
Exit Sub
Z:
    Fny = SyzSS("Sku CurRateAc VdtFm VdtTo HKD Per CA_Uom")
    Ey = Termy("Sku [     Amount] [Valid From] [Valid to] Unit per Uom")
    Into = "#IZHT086"
    T = ">ZHT086"
    Bexp = ""
    Debug.Print SqlSel_Fny_Extny_Into_Fm(Fny, Ey, Into, T, Bexp)
    Return
End Sub
Function SqlDrpCol_T_FF(T$, FF$):  SqlDrpCol_T_FF = "Alter Table [" & T & "] Drop Column " & QpFis(FF): End Function
Function SqlAddCol$(T$, QpColSpec$):    SqlAddCol = "Alter Table [" & T & "] Add Column " & QpColSpec:  End Function
Function SqlUpd$(T, SetEq$):               SqlUpd = "Update [" & T & "] set " & vbCrLf & SetEq:         End Function
