Attribute VB_Name = "MxDaoSqlFmt"
Option Compare Text
Option Explicit
Public Const SqlKw$ = "Select Update Insert Into From [Left Join] [Inner Join] Where Order Group Having"
Function SqlKwy() As String()
Static X$(): If Si(X) = 0 Then X = Termy(SqlKw)
SqlKwy = X
End Function

Private Sub FmtSql__Tst()
Dim Sql$, SqlLy$()
Dim Q: For Each Q In Itr(CQny)
    If Q = "qOHMB52" Then
        DoEvents
        Debug.Print "=================================================="
        D Q
        Sql = SqlzCQn(Q)
        SqlLy = FmtSql(Sql)
        D TabSy(SqlLy)
'        D "----"
'        D TabSy(Jnp(SqlLy))
    End If
Next
End Sub
Function FmtlSql$(Q): FmtlSql = JnCrLf(FmtSql(Q)): End Function
Function FmtSql(Q) As String(): FmtSql = FmtJnpzL(Sqpl(Q)): End Function
Private Function Sqpl$(Q) '#Sq-Phrase-lines# each line must begin with a SqKw
Dim O$: O = RplCrLf(Q)
Dim Kw: For Each Kw In SqlKwy
    O = Replace(O, Kw, vbCrLf & Kw)
Next
Sqpl = RmvPfx(O, vbCrLf)
End Function
Private Function FmtJnpzL(Sqpl$) As String()
Dim Sqpy$(): Sqpy = SplitCrLf(Sqpl)
Dim I As Bei: I = JnpBei(Sqpy)
Select Case True
Case IsEmpBei(I), I.Bix = I.Eix: FmtJnpzL = Sqpy: Exit Function
End Select
Dim F$(): F = FmtJnp(Jnp(Sqpy))
FmtJnpzL = RplAy(Sqpy, F, I.Bix)
End Function
Private Function Jnp(Sqpy$()) As String(): Jnp = AwBei(Sqpy, JnpBei(Sqpy)): End Function ' #Sql-Join-Phrases# :Ly
Private Function JnpBei(O$()) As Bei '#JoinPhrase-Bei#
Dim B%: B = JnpBix(O)
Dim E%: E = JnpEix(O, B)
JnpBei = Bei(B, E)
End Function
Private Function JnpEix%(L$(), B%) '#JoinLines-Eix
If B = 0 Then JnpEix = -1: Exit Function
Dim J%: For J = B + 1 To UB(L)
    If Not IsJnp(L(J)) Then JnpEix = J - 1: Exit Function
Next
JnpEix = J - 1
End Function
Private Function JnpBix%(L$()) '#JoinLines-Bix
Dim I: For Each I In Itr(L)
    If IsJnp(I) Then Exit Function
    JnpBix = JnpBix + 1
Next
JnpBix = 0
End Function
Private Function IsJnp(L) As Boolean
Select Case True
Case HasPfx(L, "Inner Join"), HasPfx(L, "Left Join"): IsJnp = True
End Select
End Function

'**FmtJnp
Private Sub JnpDr__Tst()
Dim Jnpln$
GoSub T1
Exit Sub
T1:
    Jnpln = "Inner Join qPHL1 ON qPHL3.PHL1 = qPHL1.PHL1"
    Ept = Sy("")
    GoTo Tst
Tst:
    Act = JnpDr(Jnpln)
    C
End Sub
Private Function FmtJnp(Jnp$()) As String(): FmtJnp = FmtStrColy(JnpStrColy(Jnp)): End Function '#Format-JoinLines-Phrases#
Private Function JnpStrColy(Jnp$()) As StrColy
Dim Dy()
Dim Ln: For Each Ln In Itr(Jnp)
    PushI Dy, JnpDr(Ln)
Next
JnpStrColy = StrColyzDy(Dy)
End Function
Private Function JnpDr(Jnpln) As String()
Dim O$: O = Jnpln
PushI JnpDr, ShfJn(O)
PushI JnpDr, ShfAlias(O)
PushI JnpDr, ShfOnLHS(O)
PushI JnpDr, "= " & Trim(RmvPfx(O, "="))
End Function
Private Function ShfJn$(OLn$): ShfJn = ShfBefSS(OLn, "As On"): End Function
Private Function ShfOnLHS$(OLn$): ShfOnLHS = ShfBefEq(OLn): End Function
Private Function ShfAlias$(OLn$): ShfAlias = Trim(RmvPfx(ShfBef(OLn, "On"), "As")): End Function
Function ShfBefSS$(OLn$, BefSS$)
Dim Bef: For Each Bef In SplitSpc(BefSS)
    ShfBefSS = ShfBefOpt(OLn, Bef)
    If ShfBefSS <> "" Then Exit Function
Next
Thw CSub, "No BefSS in OLn", "BefSS OLn", BefSS, OLn
End Function
Function ShfTermXBef$(OLn$, TermX$, Bef$)
If ShfTermX(OLn, TermX) Then Exit Function
ShfTermXBef = ShfBef(OLn, Bef)
End Function

