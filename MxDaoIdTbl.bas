Attribute VB_Name = "MxDaoIdTbl"
Option Explicit
Option Compare Text
Const CNs$ = "IdTbl"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoIdTbl."

Function LasId&(D As Database, T)
'@T ! Assume it has a field <T>Id and a "PrimaryKey", using the field as Key
ChkIsIdTbl D, T
Dim R As DAO.Recordset: Set R = D.TableDefs(T).OpenRecordset
R.Index = "PrimaryKey"
R.MoveLast
LasId = R.Fields(0).Value
End Function

Sub ChkIsIdTbl(D As Database, T, Optional Fun$ = "ChkIdTbl")
If Not IsIdTbl(D, T) Then Thw Fun, "Given table is not Id-Table (should have Id-Fld Id-Pk)", "Db T", D.Name, T
End Sub

Function IsIdTbl(D As Database, T) As Boolean
Select Case True
Case NoIdFld(D, T): Exit Function
Case NoIdPk(D, T): Exit Function
End Select
IsIdTbl = True
End Function

Function HasIdFld(D As Database, T) As Boolean
HasIdFld = D.TableDefs(T).Fields(0).Name = T & "Id"
End Function

Function NoIdFld(D As Database, T) As Boolean
NoIdFld = Not HasIdFld(D, T)
End Function

Function HasIdPk(D As Database, T) As Boolean
HasIdPk = IsEqAy(PkFny(D, T), Sy(T & "Id"))
End Function

Function NoIdPk(D As Database, T) As Boolean
NoIdPk = HasIdPk(D, T)
End Function
