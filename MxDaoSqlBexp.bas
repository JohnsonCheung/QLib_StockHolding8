Attribute VB_Name = "MxDaoSqlBexp"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoSqlBexp."
Const QQ_F_In$ = "? in (?)"
Const QQ_F_Eq$ = "? = ?"
Function Bexp_Skvy$(D As Database, T, Skvy):   Bexp_Skvy = Bexp_Fny_EqVy(SkFny(D, T), Skvy):                                 End Function
Function Bexp_E_In$(Epr$, InPrimVy):           Bexp_E_In = FmtQQ(QQ_F_In, Epr, InPrimVy):                                    End Function
Function Bexp_F_In$(F, InPrimVy):              Bexp_F_In = FmtQQ(QQ_F_In, QuoSq(F), QuoBkt(JnComma(QuoSqlPrimy(InPrimVy)))): End Function
Function Bexp_Fny_EqVy$(Fny$(), EqVy, Optional Alias$)
Dim O$()
Dim J%: For J = 0 To UB(Fny)
    PushI O, Bexp_F_Eq(QuoF(Fny(J), Alias), QuoSqlv(EqVy(J)))
Next
Bexp_Fny_EqVy = JnAnd(O)
End Function
Function Bexp_F_Eq$(F, Eqval, Optional Alias$)
Dim Fld$: Fld = QuoF(F, Alias)
Dim V$: V = QuoSqlv(Eqval)
Bexp_F_Eq = FmtQQ(QQ_F_Eq, Fld, V)
End Function

