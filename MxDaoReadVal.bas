Attribute VB_Name = "MxDaoReadVal"
Option Explicit
Option Compare Text
Const CNs$ = "DaoVal"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoReadVal."

Private Sub VzQ__Tst()
Dim D As Database
Ept = CByte(18)
Act = VzQ(D, "Select Y from [^YM]")
C
End Sub
Private Sub VzCQQ__Tst()
MsgBox VzCQ(SqlSel_FF_T("DD YY MM", "@OH", Bexp_F_Eq("YY", 20)))
MsgBox VzCQQ("Select DD from[@OH]where YY=?", 20)
End Sub
Function VzQ(D As Database, Q): VzQ = VzRs(D.OpenRecordset(Q)): End Function
Function VzCQ(Q): VzCQ = VzQ(CDb, Q): End Function
Function VzCQQ(QQSql$, ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
VzCQQ = VzQ(CDb, FmtQQAv(QQSql, Av))
End Function
Function VzQQ(D As Database, QQSql$, ParamArray Ap())
Dim Av(): Av = Ap
VzQQ = VzQ(D, FmtQQAv(QQSql, Av))
End Function
Function VzRs(A As DAO.Recordset, Optional F = 0)
If HasRec(A) Then VzRs = Nz(A.Fields(F).Value, Empty)
End Function
Function IdzSskv&(D As Database, T, Sskv):        IdzSskv = VzSskv(D, T, T & "Id", Sskv):                              End Function
Function VzSskv(D As Database, T, F, Sskv):        VzSskv = VzRs(Rs(D, SqlSel_F_T_WhF_Eq(F, T, SskFldn(D, T), Sskv))): End Function
Function VzSkvy(D As Database, T, F, Skvy()):      VzSkvy = VzQ(D, SqlSel_F_T(F, T, Bexp_Skvy(D, T, Skvy))):           End Function
Function VzF(D As Database, T, F, Optional Bexp$):    VzF = VzQ(D, SqlSel_F_T(F, T, Bexp)):                            End Function
Function VzTF(D As Database, TF$, Optional Bexp$):   VzTF = VzQ(D, SqlSel_TF(TF, Bexp)):                               End Function
Function VzArs(A As ADODB.Recordset)
If NoReczArs(A) Then Exit Function
Dim V: V = A.Fields(0).Value
If IsNull(V) Then Exit Function
VzArs = V
End Function
Function VzCnq(A As ADODB.Connection, Q): VzCnq = VzArs(A.Execute(Q)): End Function
