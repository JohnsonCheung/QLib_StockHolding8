Attribute VB_Name = "MxDaoDbInf"
Option Compare Text
Option Explicit
Const CNs$ = "Db.Inf"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDbInf."
Sub BrwCDbInf()
BrwDbInf CDb
End Sub

Sub BrwDbInf(D As Database)
BrwDsRO DbInf(D), DrsFmtozS("2000 [Brkcc TblFld Tbl]")
End Sub

Function CDbInf() As Ds
CDbInf = DbInf(CDb)
End Function

Function DbInf(D As Database) As Ds
Dim O() As Dt, T$()
T = Tny(D)
PushDt O, W1Tbl(D, T)
PushDt O, W1TblDes(D, T)
PushDt O, W1Lnk(D, T)
PushDt O, W1TblF(D, T)
PushDt O, W1Prp(D)
PushDt O, W1Fld(D, T)
With DbInf
    .DsNm = D.Name
    .DtAy = O
End With
End Function

Private Function W1TblfDr(T, Seq%, F As DAO.Field2) As Variant()
W1TblfDr = Array(T, Seq, F.Name, DtaTy(F.Type))
End Function

Private Function W1Fld(D As Database, Tny$()) As Dt
Dim Dy(), T
For Each T In Tni(D)
Next
W1Fld = DtByFF("DbFld", "Tbl Fld Pk Ty Si Dft Req Des", Dy)
End Function

Private Function W1Lnk(D As Database, Tny$()) As Dt
Dim Dy(), C$
Dim T: For Each T In Tni(D)
   C = D.TableDefs(T).Connect
   If C <> "" Then Push Dy, Array(T, C)
Next
Dim O As Dt
W1Lnk = DtByFF("DbLnk", "Tbl Connect", Dy)
End Function

Private Function W1LnkLy(D As Database) As String()
Dim T$, I
For Each I In Tny(D)
    T = I
    PushNB W1LnkLy, CnStrzT(D, T)
Next
End Function

Private Function W1Prp(D As Database) As Dt
Dim Dy()
W1Prp = DtByFF("DbPrp", "Prp Ty Val", Dy)
End Function

Private Function W1TblDes(D As Database, Tny$()) As Dt
Dim Dy(), Des$
Dim T: For Each T In Tny
    Des = TblDes(D, T)
    If Des <> "" Then
        Push Dy, Array(T, Des)
    End If
Next
W1TblDes = DtByFF("TblDes", "Tbl Des", Dy)
End Function

Private Function W1Tbl(D As Database, Tny$()) As Dt
Dim T, Dy()
For Each T In Tny
    Push Dy, Array(T, NReczT(D, T), StruT(D, T))
Next
W1Tbl = DtByFF("Tbl", "Tbl RecCnt Stru", Dy)
End Function

Private Function W1TblF(D As Database, Tny$()) As Dt
Dim Dy()
Dim T$, I
For Each I In Tni(D)
    T = I
    PushIAy Dy, W1TblfDy(D, T)
Next
W1TblF = DtByFF("TblFld", "Tbl Seq Fld Ty Si ", Dy)
End Function

Private Function W1TblfDy(D As Database, T) As Variant()
Dim F$, Seq%, I
For Each I In Fny(D, T)
    F = I
    Seq = Seq + 1
    Push W1TblfDy, W1TblfDr(T, Seq, FdzF(D, T, F))
Next
End Function

Private Sub BrwDbInf__Tst()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
BrwDbInf DutyDtaDb
End Sub

Private Sub W1Tbl__Tst()
Dim D As Database
Stop
DmpDt W1Tbl(D, Tny(D))
End Sub
