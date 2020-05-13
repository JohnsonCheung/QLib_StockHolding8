Attribute VB_Name = "MxDaoSqlFun"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoSqlFun."

Function QuoT$(T, Optional Alias$): QuoT = QuoSqln(T, Alias): End Function
Function QuoF$(F, Optional Alias$): QuoF = QuoSqln(F, Alias): End Function
Function Alias$(Alias0$): Alias = AddPfxIfNB(Alias0, "."): End Function
Function QuoSqln$(Sqln, Optional Alias0$): QuoSqln = Alias(Alias0) & W1Quo(Sqln): End Function
Private Function W1Quo$(Sqln)
If W1Shd(Sqln) Then W1Quo = QuoSq(Sqln) Else W1Quo = Sqln
End Function
Private Function W1Shd(Sqln) As Boolean: W1Shd = W1Rx.Test(Sqln): End Function
Private Function W1Rx() As RegExp
Static X As RegExp: If IsNothing(X) Then Set X = Rx("^[^A-Za-z]|[A-Za-z][^\W]+")
Set W1Rx = X
End Function
Function SqlQuoChrzT$(A As DAO.DataTypeEnum)
Const CSub$ = CMod & "SqlQuoChrzT"
Select Case A
Case _
    DAO.DataTypeEnum.dbBigInt, _
    DAO.DataTypeEnum.dbByte, _
    DAO.DataTypeEnum.dbCurrency, _
    DAO.DataTypeEnum.dbDecimal, _
    DAO.DataTypeEnum.dbDouble, _
    DAO.DataTypeEnum.dbFloat, _
    DAO.DataTypeEnum.dbInteger, _
    DAO.DataTypeEnum.dbLong, _
    DAO.DataTypeEnum.dbNumeric, _
    DAO.DataTypeEnum.dbSingle: Exit Function
Case _
    DAO.DataTypeEnum.dbChar, _
    DAO.DataTypeEnum.dbMemo, _
    DAO.DataTypeEnum.dbText: SqlQuoChrzT = "'"
Case _
    DAO.DataTypeEnum.dbDate: SqlQuoChrzT = "#"
Case Else
    Thw CSub, "Invalid DaoTy", "DaoTy", A
End Select
End Function

Function FmtEprVblAy(EprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional Sep$ = ",") As String()
Ass IsVblAy(EprVblAy)
Dim Ident%
    If IdentOpt > 0 Then
        Ident = IdentOpt
    Else
        Ident = 0
    End If
    If Ident = 0 Then
        If Pfx <> "" Then
            Ident = Len(Pfx)
        End If
    End If
Dim O$(), P$, S$, U&, J&
U = UB(EprVblAy)
Dim W%
'    W = VblWdty(EprVblAy)
For J = 0 To U
    If J = 0 Then P = Pfx Else P = ""
    If J = U Then S = "" Else S = Sep
'    Push O, VblAli(EprVblAy(J), IdentOpt:=Ident, Pfx:=P, WdtOpt:=W, Sfx:=S)
Next
FmtEprVblAy = O
End Function

Function ShdFmtSql() As Boolean
Static X As Boolean, Y As Boolean
If Not X Then X = True: Y = Cfg.Sql.FmtSql
ShdFmtSql = Y
End Function

Function FnyzPfxN(Pfx$, N%) As String()
Dim J%
For J = 1 To N
    PushI FnyzPfxN, Pfx & J
Next
End Function


Function NsetzNN(NN$) As Dictionary
Set NsetzNN = Aet(SyzSS(NN))
End Function
