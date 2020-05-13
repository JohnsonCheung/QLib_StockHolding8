Attribute VB_Name = "JMxDao"
Option Compare Text
Const CMod$ = CLib & "JMxDao."
#If False Then
Option Explicit
Function Db(Fb) As DAO.Database
Set Db = DAO.DBEngine.OpenDatabase(Fb)
End Function

Function TnyzFb(Fb) As String()
TnyzFb = TnyzDb(Db(Fb))
End Function

Function Rs(Sql$) As DAO.Recordset
Set Rs = CurrentDb.OpenRecordset(Sql)
End Function
Function RszSql(Sql$) As DAO.Recordset
Set RszSql = CurrentDb.OpenRecordset(Sql)
End Function
Function SyzRs(Rs As DAO.Recordset, Optional F = 0) As String()
With Rs
    While Not .EOF
        PushS SyzRs, .Fields(F).Value
        .MoveNext
    Wend
    .Close
End With
End Function

#End If
