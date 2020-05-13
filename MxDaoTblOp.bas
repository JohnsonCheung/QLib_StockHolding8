Attribute VB_Name = "MxDaoTblOp"
Option Compare Text
Option Explicit
Const CNs$ = "Db"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoTblOp."
Public Const C_Des$ = "Description"
Public Const TnyzMSysObjSql$ = "Select Name from MSysObjects where Type in (1,6) and Name not Like 'MSys*' and Name not Like 'f_*_Data'"
Sub DmpCFldSno(T): DmpFldSno CDb, T: End Sub
Sub DmpFldSno(D As Database, T)
Dim J%
Dim F As DAO.Field: For Each F In Td(D, T).Fields
    Debug.Print J, F.Name, F.OrdinalPosition
    J = J + 1
Next
End Sub
Sub BrwDb(D As Database): BrwFb D.Name: End Sub
Sub BrwDbzLasTmp(): BrwDb LasTmpDb: End Sub
Sub ClsDb(D As Database):
On Error Resume Next
D.Close
End Sub

Sub CrtQry(D As Database, Qn$, Sql$): D.QueryDefs.Append NwQd(Qn, Sql): End Sub
Function NwQd(Qn$, Sql$) As DAO.QueryDef
Set NwQd = New DAO.QueryDef: NwQd.Name = Qn: NwQd.Sql = Sql
End Function
Sub CrtTbl(D As Database, T, FldDclAy): D.Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy)): End Sub
Sub CrtCTmpA(): CrtTmpA CDb: End Sub
Sub CrtTmpA(D As Database)
DrpT D, "#A"
D.TableDefs.Append TmpATd
End Sub

Function PthzDb$(D As Database): PthzDb = Pth(D.Name): End Function
Function ReOpnDb(D As Database) As Database: Set ReOpnDb = Db(D.Name): End Function
Sub DmpNRec(D As Database): Dmp FmtNRec(D): End Sub

Sub DrpDbIfTmp(D As Database)
If IsTmpDb(D) Then
    Dim N$
    N = D.Name
    D.Close
    DltFfn N
End If
End Sub
Function HasT(D As Database, T) As Boolean: HasT = HasItn(ReOpnDb(D).TableDefs, T): End Function

'**Drp-Tbl
Sub DrpCT(T): DrpT CDb, T: End Sub
Sub DrpT(D As Database, T)
If HasT(D, T) Then D.Execute "Drop Table [" & T & "]"
End Sub
Sub DrpTmpCTbl(): DrpTmpTbl CDb: End Sub
Sub DrpTmpTbl(D As Database): DrpTny D, TmpTny(D): End Sub
Sub DrpTny(D As Database, Tny)
Dim T: For Each T In Itr(Tny)
    DrpT D, T
Next
End Sub
Sub DrpCTny(Tny): DrpTny CDb, Tny: End Sub
Sub DrpCTT(TT$): DrpTT CDb, TT: End Sub
Sub DrpTT(D As Database, TT$): DrpTny D, SyzSS(TT): End Sub
Sub DrpCTbAp(ParamArray TblAp())
Dim Av(): Av = TblAp
DrpTny CDb, Av
End Sub
Sub DrpTblAp(D As Database, ParamArray TblAp())
Dim Av(): Av = TblAp
DrpTny D, Av
End Sub

Function DszDb(D As Database, Optional DsNm$) As Ds
Dim Nm$
If DsNm = "" Then
    Nm = D.Name
Else
    Nm = DsNm
End If
DszDb = DszTny(D, Tny(D), Nm)
End Function

Function DszTny(D As Database, Tny$(), Optional DsNm$) As Ds
Dim T: For Each T In Tny
    PushDt DszTny.DtAy, DtzT(D, CStr(T))
Next
End Function

Sub EnsTmpTbl(D As Database)
If HasT(D, "#Tmp") Then Exit Sub
D.Execute "Create Table [#Tmp] (AA Int, BB Text 10)"
End Sub

Function FFzT$(D As Database, T): FFzT = Termln(Fny(D, T)): End Function
Function CFFzT$(T): CFFzT = FFzT(CDb, T): End Function

Function FldDesDic(D As Database) As Dictionary
Dim T$, I, J, F$, Des$
Set FldDesDic = New Dictionary
For Each I In Tni(D)
    T = I
    For Each J In Fny(D, T)
        F = J
        Des = FldDes(D, T, F)
        If Des <> "" Then FldDesDic.Add T & "." & F, D
    Next
Next
End Function

Function FmtNRec(D As Database) As String()
Dim T$(): T = Tny(D)
BfrV "Fb   " & D.Name
BfrV "NTbl " & Si(T)
Dim I, J%
For Each I In Itr(T)
    J = J + 1
    BfrV AliR(J, 3) & " " & AliR(NReczT(D, I), 7) & " " & I
Next
FmtNRec = BfrLy
End Function

Function AnyFF(D As Database, T, FF$) As Boolean
AnyFF = HasSubAy(Fny(D, T), Termy(FF))
End Function

Function HasCQn(Qn) As Boolean: HasCQn = HasQn(CDb, Qn): End Function
Function HasQn(D As Database, Qn) As Boolean
Dim W$: W = FmtQQ("Name='?' and Type=5", Qn)
HasQn = HasRecQ(D, SqlSelStar_Fm("MSysObjects", W))
End Function

Function HasRecQn(D As Database, Qn) As Boolean: HasRecQn = HasRec(RszQn(D, Qn)): End Function
Function HasRecQ(D As Database, Q) As Boolean: HasRecQ = HasRec(Rs(D, Q)): End Function
Function HasRecCQ(Q) As Boolean: HasRecCQ = HasRecQ(CDb, Q): End Function
Function HasRecCQn(Qn) As Boolean: HasRecCQn = HasRecQn(CDb, Qn): End Function
Function HasRecCTbl(T) As Boolean: HasRecCTbl = HasRecT(CDb, T): End Function
Function HasRecT(D As Database, T) As Boolean: HasRecT = HasRec(RszT(D, T)): End Function

Function HasTbl(T) As Boolean: HasTbl = HasT(CDb, T): End Function
Function HasTblByMSys(D As Database, T) As Boolean
Dim W$: W = FmtQQ("Type in (1,6) and Name='?'", T)
HasTblByMSys = HasRecQ(D, SqlSelStar_Fm("MSysObjects", W))
End Function

Function IsDbOk(D As Database) As Boolean
On Error GoTo X
IsDbOk = D.Name = D.Name
Exit Function
X:
End Function

Function IsTmpDb(D As Database) As Boolean: IsTmpDb = PthzDb(D) = TmpDbPth: End Function

'**Brw-Qry
Sub BrwCQd(): BrwQd CDb: End Sub
Sub BrwQd(D As Database): BrwS12y QryS12y(D): End Sub

Function CQryS12y() As S12(): CQryS12y = QryS12y(CDb): End Function
Function QryS12y(D As Database) As S12()
Dim Q: For Each Q In Itr(Qny(D))
    PushS12 QryS12y, S12(Q, FmtlSql(SqlzQn(D, Q)))
Next
End Function
Function CQny() As String(): CQny = Qny(CDb): End Function

Function Qny(D As Database) As String()
Qny = SyzQ(D, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'")
End Function

Function FrmNy(D As Database) As String(): FrmNy = Itn(FrmCntr(D).Documents): End Function
Function RptNy(D As Database) As String(): RptNy = Itn(RptCntr(D).Documents): End Function
Function CFrmNy() As String(): CFrmNy = FrmNy(CDb): End Function
Function CRptNy() As String(): CRptNy = RptNy(CDb): End Function

Private Function FrmCntr(D As Database) As Container: Set FrmCntr = Cntr(D, "Forms"): End Function
Private Function RptCntr(D As Database) As Container: Set RptCntr = Cntr(D, "Reports"): End Function

Private Function Cntr(D As Database, Cntrn) As Container
Dim C As Container: For Each C In D.Containers
    If C.Name = Cntrn Then Set Cntr = C
Next
End Function

Sub RunQQ(D As Database, QQ$, ParamArray Ap())
Dim Av(): Av = Ap
RunQQAv D, QQ, Av
End Sub
Sub RunQQAv(D As Database, QQ$, Av()): RunQ D, FmtQQAv(QQ, Av): End Sub 'Ret : Run the %Sql by building from &FmtQQ(@QQ,@Av) in @D

Function RszCQn(Qn) As DAO.Recordset: Set RszCQn = RszQn(CDb, Qn): End Function
Function RszQn(D As Database, Qn) As DAO.Recordset: Set RszQn = Qd(D, Qn).OpenRecordset: End Function
Function RszCQ(Q) As DAO.Recordset: Set RszCQ = CRs(Q): End Function
Function CRs(Q) As DAO.Recordset: Set CRs = Rs(CDb, Q): End Function
Function RszQ(D As Database, Q) As DAO.Recordset: Set RszQ = Rs(D, Q): End Function
Function Rs(D As Database, Q) As DAO.Recordset
Const CSub$ = CMod & "Rs"
On Error GoTo X
Set Rs = D.OpenRecordset(Q)
Exit Function
X: Thw CSub, "Error in opening Rs", "Er Sql Db", Err.Description, Q, D.Name
End Function

Function RszQQ(D As Database, QQ$, ParamArray Ap()) As DAO.Recordset
Dim Av():  Av = Ap
Set RszQQ = Rs(D, FmtQQAv(QQ, Av))
End Function

Sub SetFldDesByDi(D As Database, TFDes As Dictionary)
Dim T$, F$, Des$, TDotF$, I, J
For Each I In TFDes.Keys
    TDotF = I
    Des = TFDes(TDotF)
    If HasDot(TDotF) Then
        AsgBrkDot TDotF, T, F
        SetFldDes D, T, F, Des
    Else
        For Each J In Tny(D)
            T = J
            If HasFld(D, T, F) Then
                SetFldDes D, T, F, Des
            End If
        Next
    End If
Next
End Sub

Function SrcTny(D As Database) As String(): SrcTny = SyzItrP(D.TableDefs, "SourceTableName"): End Function

Function TdStrAy(D As Database, TT$) As String()
Dim T: For Each T In ItrzTml(TT)
    PushI TdStrAy, TdStrzT(D, T)
Next
End Function

Function TmpFbAy() As String(): TmpFbAy = Ffny(TmpDbPth, "*.accdb"): End Function
Function TmpTny(D As Database) As String(): TmpTny = AwPfx(Tny(D), "#"): End Function
Function CTni(): Asg Tni(CDb), CTni: End Function
Function Tni(D As Database): Asg Itr(Tny(D)), Tni: End Function
Function InpTbnItr(D As Database): Asg Itr(InpTny(D)), InpTbnItr: End Function
Function CTT$(): CTT = TT(CDb): End Function
Function TT$(D As Database): TT = Termln(Tny(D)): End Function
Function CTny() As String(): CTny = Tny(CDb): End Function
Function TnyzA(A As Access.Application) As String(): TnyzA = Tny(A.CurrentDb): End Function
Function LclTny(D As Database) As String(): LclTny = AePfx(ItnWhPrpBlnk(D.TableDefs, "Connect"), "MSys"): End Function
Function CnzT$(D As Database, T): CnzT = D.TableDefs(T).Connect: End Function
Function CnzCT$(T): CnzCT = CnzT(CDb, T): End Function
Function LnkdTny(D As Database) As String(): LnkdTny = ItnWhPrpBlnk(D.TableDefs, "Connect"): End Function
Function CLclTny() As String(): CLclTny = LclTny(CDb): End Function
Function CLnkdTny() As String(): CLnkdTny = LnkdTny(CDb): End Function
Function OupTny(D As Database) As String(): OupTny = AwPfx(Tny(D), "@"): End Function
Function Tny(D As Database, Optional IsReOpn As Boolean) As String()
If IsReOpn Then Set D = DAO.DBEngine.OpenDatabase(D.Name)
Dim T As TableDef: For Each T In D.TableDefs
    If Not IsTdSys(T) Then
        If Not IsTdHid(T) Then
            PushI Tny, T.Name
        End If
    End If
Next
End Function

Function Tny1(D As Database) As String()
Dim T As TableDef, O$()
Dim X As DAO.TableDefAttributeEnum
X = DAO.TableDefAttributeEnum.dbHiddenObject Or DAO.TableDefAttributeEnum.dbSystemObject
For Each T In D.TableDefs
    Select Case True
    Case T.Attributes And X
    Case Else
        PushI Tny1, T.Name
    End Select
Next
End Function

Function CTnyByADO() As String(): CTnyByADO = TnyByADO(CDb): End Function
Function TnyByADO(D As Database) As String(): TnyByADO = TnyzFb(D.Name): End Function
Function CInpTny() As String(): CInpTny = InpTny(CDb): End Function
Function InpTny(D As Database) As String(): InpTny = AwLik(Tny(D), ">*"): End Function
Function TnyzMSysObj(D As Database) As String(): TnyzMSysObj = SyzQ(D, TnyzMSysObjSql): End Function

Private Sub BrwT__Tst()
Dim D As Database
DrpTT D, "#A #B"
RunQ D, "Select Distinct Sku,BchNo,CLng(Rate) as RateRnd into [#A] from [#T]"
BrwTT D, "#A #T #B"
End Sub

Private Sub DszDb__Tst()
Dim D As Database, Tny0, Act As Ds, Ept As Ds
Stop
YY1:
    Set D = Db(DutyDtaFb)
    Act = DszDb(D)
    BrwDs Act
    Exit Sub
YY2:
    Tny0 = "Permit PermitD"
    'Set Act = Ds( Tny0)
    Stop
End Sub

Sub DrpCTblByPfxss(Pfxss$): DrpTblByPfxss CDb, Pfxss: End Sub
Sub DrpTblByPfxss(D As Database, Pfxss$): DrpTny D, AwPfxss(Tny(D), Pfxss): End Sub

Function DftDb(D As Database) As Database: Set DftDb = IIf(IsNothing(D), CDb, D): End Function
Function CDb() As Database: Set CDb = CurrentDb: End Function
Function CDbn$(): CDbn = CDb.Name: End Function
Function CFb$(): CFb = CurrentDb.Name: End Function
Function CDbPth$(): CDbPth = Pth(CurrentDb.Name): End Function

