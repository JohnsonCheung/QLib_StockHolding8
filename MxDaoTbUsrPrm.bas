Attribute VB_Name = "MxDaoTbUsrPrm"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoTbUsrPrm."
':Pmn: :Nm #Parameter-Name#
':Pmv: :Val #Parameter-Value#
':CPmv: :Val #CurDb-Parameter-Value#

Function Pmv(D As Database, Pmn$)
Pmv = VzQ(D, SelPmSql(Pmn))
End Function

Function CPmv(Pmn$)
CPmv = Pmv(CDb, Pmn)
End Function

Function PmNy(D As Database) As String()
PmNy = QSrt(FnyzRs(RszT(D, "UsrPrm")))
End Function

Sub DmpPm(D As Database)
DmpRec RszQ(D, SelAllPmSql)
End Sub

Sub DmpCPm()
DmpPm CDb
End Sub

Function CPmNy() As String()
CPmNy = PmNy(CDb)
End Function

Function CUsr$()
CUsr = "User" 'Environ$("USERNAME")
End Function

Sub SetCPmv(Pmn$, V)
SetPmv CDb, Pmn, V
End Sub

Sub SetPmv(D As Database, Pmn$, V)
UpdRsV Rs(D, SelPmSql(Pmn)), V
End Sub

Function SelPmSql$(Pmn$)
SelPmSql = FmtQQ("Select [?] from UsrPrm where Usr='?'", Pmn, CUsr)
End Function

Function SelAllPmSql$()
SelAllPmSql = FmtQQ("Select * from UsrPrm where Usr='?'", CUsr)
End Function
