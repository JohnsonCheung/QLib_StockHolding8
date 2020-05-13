Attribute VB_Name = "MxDaoPrp"
Option Compare Text
Option Explicit
Const CNs$ = "Dao.Prp"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoPrp."

Function FldDes$(D As Database, T, F$)
FldDes = FldPv(D, T, F, C_Des)
End Function

Function FldDeszTd$(A As DAO.Field)
FldDeszTd = DaoPv(A.Properties, C_Des)
End Function

Function FldPv(D As Database, T, F$, P$)
If Not HasFldPrp(D, T, F, P) Then Exit Function
FldPv = D.TableDefs(T).Fields(F).Properties(P).Value
End Function

Function HasPrp(Prps As DAO.Properties, P) As Boolean
HasPrp = HasItn(Prps, P)
End Function

Function HasTblPrp(D As Database, T, P) As Boolean
HasTblPrp = HasPrp(D.TableDefs(T).Properties, P)
End Function

Function HasFldPrp(D As Database, T, F$, P$) As Boolean
HasFldPrp = HasItn(D.TableDefs(T).Fields(F).Properties, P)
End Function

Function PrpDyzFd(A As DAO.Field) As Variant()
Dim PrpV, I, P$, V
For Each I In Itn(A.Properties)
    V = DaoPv(A, P)
    PushI PrpDyzFd, Array(P, V, TypeName(V))
Next
End Function

Function PrpNyzFd(A As DAO.Field) As String()
PrpNyzFd = Itn(A.Properties)
End Function
Sub objPrps(): End Sub
Sub ObjPrps1()
End Sub
Function PrpszObj(ObjWiPrps) As DAO.Properties
Const CSub$ = CMod & "PrpszObj"
On Error GoTo X
Set PrpszObj = ObjWiPrps.Properties
Exit Function
X:
    Dim E$: E = Err.Description
    Thw CSub, "Obj does not have prp-[Properties]", "Obj-Tyn Er", TypeName(ObjWiPrps), E
End Function

Sub SetFldDes(D As Database, T, F$, Des$)
SetFldPv D, T, F, C_Des, Des
End Sub

Sub SetTblPv(D As Database, T, P$, V)
Dim Td As DAO.TableDef: Set Td = D.TableDefs(T)
If IfSetPrpsPv(Td.Properties, P, V) Then Exit Sub
Td.Properties.Append Td.CreateProperty(P, DaoTy(V), V)
End Sub
Sub SetCTblPv(T, P$, V)
SetTblPv CDb, T, P, V
End Sub
Sub SetCTblDes(T, Optional Des$)
SetTblDes CDb, T, Des
End Sub
Sub SetTblDes(D As Database, T, Optional Des$)
SetTblPv D, T, "Description", Des
End Sub

Sub DltPrp(D As Database, P)
Dim Prps As DAO.Properties: Set Prps = D.Properties
If HasPrp(Prps, P) Then
    Prps.Delete P
End If
End Sub
Function CTblDes$(T)
CTblDes = TblDes(CDb, T)
End Function

Function TblDes$(D As Database, T)
TblDes = TblPrp(D, T, "Description")
End Function

Function TblPrp(D As Database, T, P)
If Not HasTblPrp(D, T, P) Then Exit Function
TblPrp = D.TableDefs(T).Properties(P).Value
End Function

Private Sub SetTbPv__Tst()
Dim D As Database: Set D = TmpDb
DrpT D, "Tmp"
RunQ D, "Create Table Tmp (F1 Text)"
SetTblPv D, "Tmp", "XX", "AFdf"
Debug.Assert TblPrp(D, "Tmp", "XX") = "AFdf"
End Sub

Private Sub FldPv__Tst()
Dim P$, Db As Database, T, F$, V
GoSub T0
Exit Sub
T0:
    Set Db = TmpDb
    RunQ Db, "Create Table Tmp (AA Text)"
    T = "Tmp"
    F = "AA"
    P = "Ele"
    V = "Ele1234"
    GoTo Tst
Tst:
    FldPv(Db, T, F, P) = V
    Ass FldPv(Db, T, F, P) = V
    Dim Fd As DAO.Field: Set Fd = FdzF(Db, T, F)
    Stop
    DmpDy PrpDyzFd(Fd)
    Return
End Sub

Private Sub PrpDyzFd__Tst()
Dim Db As Database: Set Db = DutyDtaDb
Dim Fd As DAO.Field
Dim Rs As DAO.Recordset
Set Rs = RszT(Db, "Permit")
Set Fd = Rs.Fields("Permit")
Debug.Print Fd.Value
DmpDy PrpDyzFd(Fd)
End Sub

Private Sub PrpNy__Tst()
Dim Db As Database: Set Db = DutyDtaDb
Dim Fd As DAO.Field
Set Fd = FdzF(Db, "Permit", "Permit")
D PrpNyzFd(Fd)
End Sub

Sub SetFldPv(D As Database, T, F, P, V)
Dim Tdef As DAO.TableDef: Set Tdef = Td(D, T)
Dim Fdef As DAO.Field: Set Fdef = Tdef.Fields(F)
If IfSetPrpsPv(Fdef.Properties, P, V) Then Exit Sub
Fdef.Properties.Append Tdef.CreateProperty(P, DaoTy(V), V) ' will break if V=""
End Sub

Private Function IfSetPrpsPv(Ps As DAO.Properties, P, V) As Boolean
If HasItn(Ps, P) Then
    Ps(P).Value = V
    IfSetPrpsPv = True
End If
End Function


Function CvPrps(DaoObjWiPrps) As DAO.Properties
On Error GoTo X
Set CvPrps = DaoObjWiPrps.Properties
Exit Function
X:
    Dim E$: E = Err.Description
    Thw "CvPrps", "Calling DaoObjWiPrps.Properties error", "TypeName-DaoObjWiPrps Er", TypeName(DaoObjWiPrps), E
End Function
Sub DltQryFldPrp(Q, F, P)
End Sub
Sub SetQryFldPrp(Q, F, P, V)
Dim Qd As QueryDef: Set Qd = CurrentDb.QueryDefs(Q)
Dim Fds As DAO.Fields: Set Fds = Qd.Fields
Dim Fd As DAO.Field: Set Fd = Fds(F)
If HasItn(F.Properties, P) Then
    Fd.Properties(P).Value = V
Else
    Fd.Properties.Append Fd.CreateProperty(P, DaoTy(V), V)
End If
End Sub
