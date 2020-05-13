Attribute VB_Name = "MxDaoDbTblInf"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDbTblInf."
Const Skn$ = "SecondaryKey"

Function COrdinalPosition%(T, F): COrdinalPosition = OrdinalPosition(CDb, T, F): End Function
Function OrdinalPosition%(D As Database, T, F): OrdinalPosition = D.TableDefs(T).Fields(F).OrdinalPosition: End Function
Function FmtOrdinalPosition(D As Database, T) As String(): FmtOrdinalPosition = FmtDrsR(OrdinalPositionDrs(D, T)): End Function ' Return the formatted-string-array of each fields ordinal-position of @T in @D.  The format is [Ix Fldn OrdinalPosition]
Private Function OrdinalPositionDrs(D As Database, T) As Drs: OrdinalPositionDrs = Drs(SyzSS("OrdPos Fldn"), OrdinalPositionDy(D, T)): End Function
Private Function OrdinalPositionDy(D As Database, T) As Variant()
Dim J%: Dim F As DAO.Field: For Each F In D.TableDefs(T).Fields
    PushI OrdinalPositionDy, Array(F.OrdinalPosition, F.Name)
    J = J + 1
Next
End Function

Sub DmpCOrdinalPosition(T): DmpOrdinalPosition CDb, T: End Sub
Sub DmpOrdinalPosition(D As Database, T): DmpAy FmtOrdinalPosition(D, T): End Sub

Sub AddFld(D As Database, T, F$, Ty As DataTypeEnum, Optional Si%, Optional Precious%)
If HasFld(D, T, F) Then Exit Sub
Dim S$, SqSpect$
SqSpect = sqlTyzDao(Ty, Si, Precious)
S = FmtQQ("Alter Table [?] Add Column [?] ?", T, F, Ty)
D.Execute S
End Sub

Sub AddFdzEpr(D As Database, T, F$, Epr$, Ty As DAO.DataTypeEnum): D.TableDefs(T).Fields.Append Fd(F, Ty, Epr:=Epr): End Sub

Sub AsgColApzDrsFF(D As Drs, FF$, ParamArray OColAp())
Dim F, J%
For Each F In Termy(FF)
    OColAp(J) = ColzDrs(D, CStr(F))
    J = J + 1
Next
End Sub

Sub BrwTblzByDt(D As Database, T): BrwDt DtzT(D, T): End Sub
Function CnStrzT$(D As Database, T): CnStrzT = D.TableDefs(T).Connect: End Function
Function AetzF(D As Database, T, F$) As Dictionary: Set AetzF = AetzRs(RszF(D, T, F)): End Function
Function AetzTF(D As Database, TF$) As Dictionary: Set AetzTF = AetzRs(RszTF(D, TF)): End Function
Function RszF(D As Database, T, F$) As DAO.Recordset: Set RszF = Rs(D, SqlSel_F_T(F, T)): End Function
Sub CrtEmpTblzDrs(D As Database, T, Drs As Drs): CrtTblzShtTyscfBql D, T, ShtTyscfBqlzDrs(Drs): End Sub
Function CsyzCT(T) As String(): CsyzCT = CsyzT(CDb, T): End Function
Function CsyzT(D As Database, T) As String(): CsyzT = CsyzRs(RszT(D, T)): End Function

Function DaoTyzNumCol$(NumCol)
ChkIsAy NumCol, CSub
Dim O As VbVarType: O = VarType(NumCol(0))
If Not IsNumzVbTy(O) Then Stop
Dim V: For Each V In NumCol
    O = MaxNumVbTy(O, VarType(V))
Next
DaoTyzNumCol = DaoTyzVb(O)
End Function

Function DaoTyzF(D As Database, T, F) As DAO.DataTypeEnum: DaoTyzF = D.TableDefs(T).Fields(F).Type: End Function
Function DaoTyzCF(T, F) As DAO.DataTypeEnum:  DaoTyzCF = DaoTyzF(CDb, T, F): End Function
Function DaoTyzCTF(TF$) As DAO.DataTypeEnum: DaoTyzCTF = DaoTyzTF(CDb, TF):  End Function
Function BrkTF(TF$) As S12
Dim A As S12: A = BrkDot(TF)
BrkTF = S12(RmvSqBkt(A.S1), RmvSqBkt(A.S2))
End Function
Function DaoTyzTF(D As Database, TF$) As DAO.DataTypeEnum
Dim A As S12: A = BrkTF(TF)
DaoTyzTF = DaoTyzF(D, A.S1, A.S2)
End Function

Function CntDizTF(D As Database, T, F$) As Dictionary: Set CntDizTF = CntDizRs(RszF(D, T, F$)): End Function

Sub DrpFF(D As Database, T, FF$)
Dim F$, I
For Each I In TermItr(FF)
    F = I
    D.Execute SqlDrpCol_T_F(T, F)
Next
End Sub

Function DrszT(D As Database, T) As Drs: DrszT = DrszRs(RszT(D, T)): End Function
Function DtzT(D As Database, T) As Dt: DtzT = Dt(T, Fny(D, T), DyzT(D, T)): End Function
Function DyzT(D As Database, T) As Variant(): DyzT = DyzRs(RszT(D, T)): End Function
Function DyzTFF(D As Database, T, FF$) As Variant(): DyzTFF = DyzQ(D, SqlSel_FF_T(FF, T)): End Function
Function Fds(D As Database, T) As DAO.Fields: Set Fds = D.TableDefs(T).OpenRecordset.Fields: End Function
Function CFny(T) As String(): CFny = Fny(CDb, T): End Function
Function Fny(D As Database, T) As String(): Fny = Itn(ReOpnDb(D).TableDefs(T).Fields): End Function
Function FnyzT(D As Database, T) As String(): FnyzT = Fny(D, T): End Function
Function FstUniqIdx(D As Database, T) As DAO.Index: Set FstUniqIdx = FstObjByTruePrp(D.TableDefs(T).Indexes, "Unique"): End Function
Function HasFld(D As Database, T, F$) As Boolean: HasFld = HasItn(D.TableDefs(T).Fields, F): End Function
Function HasIdx(D As Database, T, Idxn$) As Boolean: HasIdx = HasItn(D.TableDefs(T).Indexes, Idxn): End Function
Function NoPk(D As Database, T) As Boolean: NoPk = HasPk(D, T): End Function
Function HasNrmPkzTd(A As DAO.TableDef) As Boolean: HasNrmPkzTd = HasTruePrp(A.Indexes, "Primary"): End Function
Function HasSk(D As Database, T) As Boolean: HasSk = Not IsNothing(SkIdx(D, T)): End Function
Function HasNrmPk(D As Database, T) As Boolean: HasNrmPk = HasNrmPkzTd(D.TableDefs(T)): End Function



Function HasId(D As Database, T, Id&) As Boolean
If HasPk(D, T) Then HasId = HasRec(RszId(D, T, Id))
End Function

Function HasPk(D As Database, T) As Boolean
Dim Pk$(): Pk = PkFny(D, T)
If Si(Pk) <> 1 Then Exit Function
HasPk = Pk(0) = T & "Id"
End Function

Function HasPkzTd(A As DAO.TableDef) As Boolean
If Not HasPkzTd(A) Then Exit Function
Dim Pk$(): Pk = PkFnyzTd(A): If Si(Pk) <> 1 Then Exit Function
Dim P$: P = A.Name & "Id"
If Pk(0) <> P Then Exit Function
HasPkzTd = A.Fields(0).Name <> P
End Function

Function HasSkzTd(A As DAO.TableDef) As Boolean
If HasItn(A.Indexes, Skn) Then HasSkzTd = A.Indexes(Skn).Unique
End Function

Sub AddIdxzF(D As Database, T, Idxn$, F): AddIdxzFny D, T, Idxn, Sy(F): End Sub

Sub AddIdxzFny(D As Database, T, Idxn$, Fny$())
Dim TT As DAO.TableDef, I As DAO.Index
Set TT = Td(D, T)
Set I = T
TT.Indexes.Append NwIdxzTd(TT, Idxn, Fny)
TT.CreateIndex (Idxn)
End Sub

Function NwIdxzTd(Td As DAO.TableDef, Idxn$, Fny$()) As DAO.Index: Td.CreateIndex Idxn: End Function
Function Idx(D As Database, T, N) As DAO.Index: Set Idx = IdxzTd(Td(D, T), N): End Function
Function IdxzTd(Td As DAO.TableDef, N) As DAO.Index: Set IdxzTd = Td.Indexes(N): End Function
Function IntAyzF(D As Database, T, F$) As Integer(): IntAyzF = IntAyzQ(D, FmtQQ("Select [?] from [?]", F, T)): End Function
Function IsHidTbl(D As Database, T) As Boolean: IsHidTbl = (D.TableDefs(T).Attributes And DAO.TableDefAttributeEnum.dbHiddenObject) <> 0: End Function
Function IsLnk(D As Database, T) As Boolean: IsLnk = IsLnkzFb(D, T) Or IsLnkzFx(D, T): End Function
Function IsLnkzFb(D As Database, T) As Boolean: IsLnkzFb = HasPfx(CnStrzT(D, T), ";Database="): End Function
Function IsLnkzFx(D As Database, T) As Boolean: IsLnkzFx = HasPfx(CnStrzT(D, T), "Excel"): End Function

Function IsMemCol(Col) As Boolean
Dim I: For Each I In Col
    If IsStr(I) Then
        If Len(I) > 255 Then IsMemCol = True: Exit Function
    End If
Next
End Function

Function IsNumzVbTy(A As VbVarType) As Boolean
Select Case A
Case vbByte, vbInteger, vbLong, vbSingle, vbDecimal, vbDouble, vbCurrency: IsNumzVbTy = True
End Select
End Function

Function IsSysTbl(D As Database, T) As Boolean: IsSysTbl = (D.TableDefs(T).Attributes And DAO.TableDefAttributeEnum.dbSystemObject) <> 0: End Function
Function IxzF%(D As Database, T, F)
Dim O%, I As DAO.Field: For Each I In D.TableDefs(T).Fields
    If I.Name = F Then IxzF = O: Exit Function
    O = O + 1
Next
IxzF = -1
End Function

Function JnQSqCommaSpcAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnQSqCommaSpcAp = JnQSqCommaSpc(SyzAy(Av))
End Function

Sub KillIfTmpDb(D As Database)
If IsTmpDb(D) Then
    Dim Fb$: Fb = D.Name
    ClsDb D
    Kill Fb
End If
End Sub

Function LasUpdTim(D As Database, T) As Date: LasUpdTim = TblPrp(D, T, "LastUpdated"): End Function

Function Lnkinf(D As Database) As String()
Dim T: For Each T In Tni(D)
    PushI Lnkinf, LnkinfzT(D, T)
Next
End Function

Function LnkinfzT$(D As Database, T)
Dim O$, LnkFx$, LnkW$, LnkFb$, LnkT$
Select Case True
Case IsLnkzFx(D, T): LnkinfzT = FmtQQ("LnkFx(?).LnkFxw(?).Tbl(?).Db(?)", CnStrzDbt(D, T), SrcTn(D, T), T, D.Name)
Case IsLnkzFb(D, T): LnkinfzT = FmtQQ("LnkFb(?).LnkTbl(?).Tbl(?).Db(?)", CnStrzDbt(D, T), SrcTn(D, T), T, D.Name)
End Select
End Function

Function LoflzDbt$(D As Database, T): LoflzDbt = TblPrp(D, T, "Lofl"): End Function

Function MaxNumVbTy(A As VbVarType, B As VbVarType) As VbVarType
Const CSub$ = CMod & "MaxNumVbTy"
If A = B Then MaxNumVbTy = A: Exit Function
If Not IsNumzVbTy(B) Then Thw CSub, "Given B is not NumVbTy", "B-VarType", B
Dim O As VbVarType
Select Case A
Case VbVarType.vbByte:      O = B
Case VbVarType.vbInteger:   O = IIf(B = vbByte, A, B)
Case VbVarType.vbLong:      O = IIf((B = vbByte) Or (B = vbInteger), A, B)
Case VbVarType.vbSingle:    O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong), A, B)
Case VbVarType.vbDecimal:   O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle), A, B)
Case VbVarType.vbDouble:    O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle) Or (B = vbDecimal), A, B)
Case VbVarType.vbCurrency:  O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle) Or (B = vbDecimal) Or (B = vbDouble), A, B)
Case Else:                  Thw CSub, "Given A is not NumVbTy", "A-VarType", A
End Select
MaxNumVbTy = O
End Function

Function NColzT&(D As Database, T): NColzT = D.TableDefs(T).Fields.Count: End Function
Function NReczFxw&(Fx, Wsn, Optional Bexp$): NReczFxw = VzCnq(CnzFx(Fx), SqlSelCnt_Fm(AxTbn(Wsn), Bexp)): End Function
Function NReczT&(D As Database, T, Optional Bexp$): NReczT = VzQ(D, SqlSelCnt_Fm(T, Bexp)): End Function
Function CNReczT&(T, Optional Bexp$): CNReczT = NReczT(CDb, T, Bexp): End Function
Function NxtId&(D As Database, T): NxtId = VzQ(D, FmtQQ("select Max(?Id) from [?]", T, T)) + 1: End Function
Function PkFny(D As Database, T) As String(): PkFny = FnyzIdx(PkIdx(D, T)): End Function
Function PkFnyT(T) As String(): PkFnyT = PkFny(CurrentDb, T): End Function
Function PkFnyzTd(A As DAO.TableDef) As String(): PkFnyzTd = FnyzIdx(PkizTd(A)): End Function
Function PkIdx(D As Database, T) As DAO.Index: Set PkIdx = PkizTd(D.TableDefs(T)): End Function
Function PkIdxn$(D As Database, T): PkIdxn = Objn(PkIdx(D, T)): End Function
Function PkizTd(A As DAO.TableDef) As DAO.Index: Set PkizTd = FstObjByNm(A.Indexes, Pkn): End Function

Sub RenFlds(D As Database, T, FmFF$, ToFF$)
Dim FmFny$(): FmFny = FnyzFF(FmFF)
Dim ToFny$(): ToFny = FnyzFF(ToFF)
Dim J%: For J = UBound(FmFny) To 0 Step -1
    RenFld D, T, FmFny(J), ToFny(J)
Next
End Sub
Sub RenCFlds(T, FmFF$, ToFF$): RenFlds CDb, T, FmFF$, ToFF$: End Sub
Sub RenFld(D As Database, T, F$, ToFld$): D.TableDefs(T).Fields(F).Name = ToFld: End Sub
Sub RenTblzAddPfx(D As Database, T, Pfx$): RenTbl D, T, Pfx & T: End Sub
Function RszId(D As Database, T, Id&) As DAO.Recordset: Set RszId = Rs(D, SqlSel_T_WhId(T, Id)): End Function
Function RszT(D As Database, T) As DAO.Recordset: Set RszT = Rs(D, SqlSelStar_Fm(T)): End Function
Function RszCT(T) As DAO.Recordset: Set RszCT = RszT(CDb, T): End Function
Function RszTF(D As Database, TF$) As DAO.Recordset: Set RszTF = D.OpenRecordset(SqlSel_TF(TF)): End Function
Function RszTFF(D As Database, T, FF$) As DAO.Recordset: Set RszTFF = RszTFny(D, T, Ny(FF)): End Function
Function RszTFny(D As Database, T, Fny$()) As DAO.Recordset: Set RszTFny = D.OpenRecordset(SqlSel_Fny_T(Fny, T)): End Function
Sub SetLoflzDbt(D As Database, T, Lofl$): SetTblPv D, T, "Lofl", Lofl: End Sub
Function ShtTyszCol$(Col())
Const CSub$ = CMod & "ShtTyszCol"
Dim O$, I
I = Itr(Col)
Select Case True
Case IsBoolItr(I): O = "B"
Case IsDteItr(I): O = "Dte"
Case IsNumItr(I): O = ShtTyzNumCol(Col)
Case IsItrStr(I): O = IIf(IsMemCol(Col), "M", "")
Case Else: Thw CSub, "Col cannot determine its type: Not [Str* Num* Bool* Dte*:Col]", "Col", Col
End Select
ShtTyszCol = O
End Function

Function ShtTyzNumCol$(Col)
ShtTyzNumCol = ShtDaoTy(DaoTyzNumCol(Col))
End Function

Function ShtTyzNumVbTy$(NumVbTy As VbVarType)
Const CSub$ = CMod & "ShtTyzNumVbTy"
Dim O$
Select Case NumVbTy
Case VbVarType.vbByte:      O = "Byt:"
Case VbVarType.vbCurrency:  O = "C:"
Case VbVarType.vbDecimal:   O = "Dec:"
Case VbVarType.vbDouble:    O = "D:"
Case VbVarType.vbInteger:   O = "I:"
Case VbVarType.vbLong:      O = "L:"
Case VbVarType.vbSingle:    O = "S:"
Case Else: Thw CSub, "NumVbTy is not numeric VbTy", "NumVbTyp", ShtTyzNumVbTy(NumVbTy)
End Select
End Function

Function ShtTyzF$(D As Database, T, F$): ShtTyzF = ShtDaoTy(DaoTyzF(D, T, F$)): End Function
Function SqzT(D As Database, T, Optional ExlFldNm As Boolean) As Variant(): SqzT = SqzRs(RszT(D, T), ExlFldNm): End Function
Function SrcFbzT$(D As Database, T): SrcFbzT = IsBet(D.TableDefs(T).Connect, "Database=", ";"): End Function
Function SrcTn$(D As Database, T): SrcTn = D.TableDefs(T).SourceTableName: End Function


Private Sub CrtDupKeyTbl__Tst()
Dim D As Database: Set D = TmpDb
DrpTT D, "#A #B"
'T = "AA"
CrtTblzDup D, "#A", "#B", "Sku BchNo"
DrpDbIfTmp D
End Sub

Private Sub CrtTblzDrs__Tst()
Dim D As Database
GoSub Z
Exit Sub
Z:
    Set D = TmpDb
    DrpTmpTbl D
    CrtTblzDrs D, "#D", SampDrs
    BrwDb D
    Return
End Sub

Private Sub PkFny__Tst()
Z:
    Dim D As Database
    Set D = Db(DutyDtaFb)
    Dim Dr(), Dy(), T, I
    For Each I In Tny(D)
        T = I
        Erase Dr
        Push Dr, T
        PushIAy Dr, PkFny(D, T)
        PushI Dy, Dr
    Next
    BrwDy Dy
    Exit Sub
End Sub

Private Sub ShtTyscfBqlzDrs__Tst()
Dim Drs As Drs
GoSub T0
Exit Sub
T0:
    Drs = SampDrs
    Ept = "A`B:B`Byt:C`I:D`L:E`D:G`S:H`C:I`Dte:J`M:K"
    GoTo Tst
Tst:
    Act = ShtTyscfBqlzDrs(Drs)
    C
    Return
End Sub
