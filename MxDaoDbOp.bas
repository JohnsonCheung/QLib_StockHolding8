Attribute VB_Name = "MxDaoDbOp"
Option Compare Text
Option Explicit
Const CNs$ = "Db.Op"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDbOp."

'**Run-Sql
Sub RunCQ(Q)
On Error GoTo X
DoCmd.RunSql Q
Exit Sub
X:
    Dim T$: T = TmpNm("#Qry_")
    CrtQry CDb, T, Q
    Thw "RunCQ", "RunCQ-RunTimEr", "Er Sql Db TmpQryNm", Err.Description, Q, CDbn, T
End Sub
Sub RunQ(D As Database, Q)
On Error GoTo X
D.Execute Q
Exit Sub
X:
    Dim T$: T = TmpNm("#Qry_")
    CrtQry D, T, Q
    Thw "RunQ", "RunQ-RunTimEr", "Er Sql Db TmpQryNm", Err.Description, Q, D.Name, T
End Sub
Sub RunSqy(D As Database, Sqy$())
Dim Q: For Each Q In Sqy
    RunQ D, Q
Next
End Sub

'**Crt-Qry
Function NwQd(N$, Sql) As DAO.QueryDef
Dim O As New QueryDef
O.Name = N
O.Sql = Sql
Set NwQd = O
End Function
Sub CrtQry(D As Database, N$, Sql): D.QueryDefs.Append NwQd(N, Sql): End Sub
Function TmpQry$(D As Database, Sql, Optional QryNmPfx$ = "#Q")
Dim N$: N = TmpNm(QryNmPfx)
CrtQry D, N, Sql
TmpQry = N
End Function
