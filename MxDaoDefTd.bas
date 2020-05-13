Attribute VB_Name = "MxDaoDefTd"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CNs$ = "Def"
Const CMod$ = CLib & "MxDaoDefTd."
Type Tdr: Td As DAO.TableDef: End Type
Function TdrzT(T) As Tdr
Set TdrzT.Td = CTd(T)
End Function
Function Tdr(A As DAO.TableDef) As Tdr
End Function
Sub AddFdzId(A As DAO.TableDef)
A.Fields.Append FdzId(A.Name)
End Sub

Sub AddFdAy(T As DAO.TableDef, FdAy() As DAO.Field)
Dim F: For Each F In FdAy
    T.Fields.Append F
Next
End Sub
Function FdAyzTy(FF$, T As DAO.DataTypeEnum) As DAO.Field()
Dim F: For Each F In FnyzFF(FF)
    PushObj FdAyzTy, Fd(F, T)
Next
End Function

Sub AddFdzLng(A As DAO.TableDef, FF$)
AddFdAy A, FdAyzTy(FF, dbLong)
End Sub

Sub AddFdzTimstmp(A As DAO.TableDef, F$)
A.Fields.Append Fd(F, DAO.dbDate, Dft:="Now")
End Sub

Sub AddFdzTxt(A As DAO.TableDef, FF$, Optional Req As Boolean, Optional Si As Byte = 255)
Dim F$, I
For Each I In Termy(FF)
    F = I
    A.Fields.Append Fd(F, dbText, Req, Si)
Next
End Sub
Function CvFds(A) As DAO.Fields
Set CvFds = A
End Function

Function CvTd(A) As DAO.TableDef
Set CvTd = A
End Function

Sub DmpTdAy(TdAy() As DAO.TableDef)
Dim I
For Each I In TdAy
    D "------------------------"
    D TdStru(I)
Next
End Sub

Function Fdy(FF$, T As DAO.DataTypeEnum) As DAO.Field2()
Dim I, F$
For Each I In Termy(FF)
    F = I
    PushObj Fdy, Fd(F, T)
Next
End Function

Function FnyzTd(A As DAO.TableDef) As String()
FnyzTd = Itn(A.Fields)
End Function

Function FnyzTdLy(TdStru$()) As String()
Dim O$(), TdStr$, I
For Each I In Itr(TdStru)
    TdStr = I
'    PushIAy O, FnyzTdLy(TdStr)
Next
FnyzTdLy = CvSy(AwDis(O))
End Function

Function IsTdEq(A As DAO.TableDef, B As DAO.TableDef) As Boolean
With A
Select Case True
Case .Name <> B.Name
Case .Attributes <> B.Attributes
Case Not IsEqIdxs(.Indexes, B.Indexes)
'Case Not FdsIsEq(.Fields, B.Fields)
Case Else: IsTdEq = True
End Select
End With
End Function

Function IsTdHid(A As DAO.TableDef) As Boolean
IsTdHid = (A.Attributes And DAO.TableDefAttributeEnum.dbHiddenObject) <> 0
End Function

Function IsTdSys(A As DAO.TableDef) As Boolean
IsTdSys = (A.Attributes And DAO.TableDefAttributeEnum.dbSystemObject) <> 0
End Function

Function IsTdLnk(A As DAO.TableDef) As Boolean
IsTdLnk = (A.Attributes And DAO.TableDefAttributeEnum.dbAttachedTable) <> 0
End Function

Function SkFnyzTdLin(TdLin) As String()
Dim A1$, T$, Rst$
    A1 = Bef(TdLin, "|")
    If A1 = "" Then Exit Function
AsgTRst A1, T, Rst
T = RmvSfx(T, "*")
Rst = Replace(Rst, "*", T)
SkFnyzTdLin = SyzSS(Rst)
End Function

Function TdStru(Td) As String()
Dim O$(), A As DAO.TableDef
Set A = Td
PushI TdStru, TdStr(A)
Dim F As DAO.Field
For Each F In A.Fields
    PushI TdStru, FdStr(F)
Next
End Function

Function TdStruzDb(D As Database) As String()
Dim T
For Each T In Tni(D)
    PushIAy TdStruzDb, TdStru(D.TableDefs(T))
Next
End Function

Function TdStruzT(D As Database, T) As String()
TdStruzT = TdStru(D.TableDefs(T))
End Function


Function TdzTFdAy(T, FdAy() As DAO.Field) As DAO.TableDef
Dim O As New TableDef
O.Name = T
Dim F: For Each F In FdAy
    O.Fields.Append F
Next
Set TdzTFdAy = O
End Function

Sub ChkTdEq(A As DAO.TableDef, B As DAO.TableDef)
Const CSub$ = CMod & "ChkTdEq"
Dim A1$(): A1 = TdStru(A)
Dim B1$(): B1 = TdStru(B)
If Not IsEqAy(A, B) Then Thw CSub, "Two 2 Td as diff", "Td-A Td-B", TdStru(A), TdStru(B)
End Sub

Property Get TmpATd() As DAO.TableDef
Dim FdAy() As DAO.Field
PushObj FdAy, FdzTxt("F1")
Set TmpATd = TdzTFdAy("Tmp", FdAy)
End Property

Sub AddPk(A As DAO.TableDef)
Const CSub$ = CMod & "AddPk"
'Any Pk Fields in A.Fields?, if no exit sub
Dim F As DAO.Field2, IdFldNm$, J%
IdFldNm = A.Name & "Id"
If IsIdFd(A.Fields(0), A.Name) Then
    A.Indexes.Append NwCPkIdx(A.Name)
    Exit Sub
End If
For J = 2 To A.Fields.Count
    If A.Fields(J).Name = IdFldNm Then Thw CSub, "The Table Id fields must be the fst fld", "I-th", J
Next
End Sub

Sub AddSk(A As DAO.TableDef, Skff$)
Dim SkFny$(): SkFny = Termy(Skff): If Si(SkFny) = 0 Then Exit Sub
A.Indexes.Append NwSkIdx(A, SkFny)
End Sub

Function CvIdxFds(A) As DAO.IndexFields
Set CvIdxFds = A
End Function

Function IsIdFd(A As DAO.Field2, T) As Boolean
If A.Name <> T & "Id" Then Exit Function
If A.Attributes <> DAO.FieldAttributeEnum.dbAutoIncrField Then Exit Function
If A.Type <> dbLong Then Exit Function
IsIdFd = True
End Function

Function NwSkIdx(T As DAO.TableDef, SkFny$()) As DAO.Index
Const CSub$ = CMod & "NewSkIdx"
Dim O As New DAO.Index
O.Name = "SecondaryKey"
O.Unique = True
If Not HasEleAy(FnyzTd(T), SkFny) Then
    Thw CSub, "Given Td does not contain all given-SkFny", "Missing-SkFny Td-Name Td-Fny Given-SkFny", T.Name & "Id", MinusAy(SkFny, FnyzTd(T)), T.Name, FnyzTd(T), SkFny
End If
Dim IdxFds As DAO.IndexFields, I
Set IdxFds = CvIdxFds(O.Fields)
For Each I In SkFny
    IdxFds.Append Fd(CStr(I))
Next
Set NwSkIdx = O
End Function
Function NwCPkIdx(T) As DAO.Index
Set NwCPkIdx = NwPkIdx(CDb, T)
End Function

Function NwPkIdx(D As Database, T) As DAO.Index
Dim O As New DAO.Index
O.Name = "PrimaryKey"
O.Primary = True
CvIdxFds(O.Fields).Append FdzId(T & "Id")
Set NwPkIdx = O
End Function

Function Qd(D As Database, Qn) As DAO.QueryDef
Set Qd = D.QueryDefs(Qn)
End Function

Function SqlzCQn$(Qn):              SqlzCQn = SqlzQn(CDb, Qn): End Function
Function SqlzQn$(D As Database, Qn): SqlzQn = Qd(CDb, Qn).Sql: End Function

Function FdzF(D As Database, T, F) As DAO.Field
Set FdzF = D.TableDefs(T).Fields(F)
End Function

Function Td(D As Database, T) As DAO.TableDef
Set Td = D.TableDefs(T)
End Function

Function CTd(T) As DAO.TableDef
Set CTd = Td(CDb, T)
End Function

Sub AppTdAy(D As Database, TdAy() As DAO.TableDef)
Dim T
For Each T In Itr(TdAy)
    D.TableDefs.Append T
Next
End Sub

