Attribute VB_Name = "MxXls"
Option Compare Text
Option Explicit
Const CNs$ = "Xls"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXls."
Public Const MaxCno% = 16384
Public Const MaxRno& = 1048576
Public Const FexeXls$ = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
':LoZZ$ = "z when used in Nm, it has special meaning.  It can occur in Cml las-one, las-snd, las-thrid chr, else it is er."
':NmBrk$ = "NmBrk is z or zx or zxx where z is letter-z and x is lowcase or digit.  NmBrk must be sfx of a cml."
':NmBrk_za$ = "It means `and`."
Enum eWsdAddWay ' Adding data to ws way
    eWcWay
    eSqWay
End Enum

Function RCzA1(R As Range) As RC
With RCzA1
    .R = R.Row
    .C = R.Column
End With
End Function

Function WbzWs(A As Worksheet) As Workbook
Set WbzWs = A.Parent
End Function

Function MainWs(A As Workbook) As Worksheet
Set MainWs = WszCdNm(A, "WsOMain")
End Function

Function WnyzRg(A As Range) As String()
WnyzRg = Wny(WbzRg(A))
End Function

Function WnyzWb(A As Workbook) As String()
WnyzWb = Itn(A.Sheets)
End Function

Private Sub XlszG__Tst()
Debug.Print XlszG.Name
End Sub

Function XlszG() As Excel.Application
'Set XlszGetObj = GetObject(FexeXls)
Set XlszG = GetObject(, "Excel.Application")
End Function

Function FstWb() As Workbook
Set FstWb = FstWbzX(Xls)
End Function

Function FstWbzX(X As Excel.Application) As Workbook
Set FstWbzX = X.Workbooks(1)
End Function

Function HasAddinFn(A As Excel.Application, AddinFn$) As Boolean
HasAddinFn = HasItn(A.AddIns, AddinFn)
End Function

Function DftXls(A As Excel.Application) As Excel.Application
If IsNothing(A) Then
    Set DftXls = NwXls
Else
    Set DftXls = A
End If
End Function

Function NRowOfLo&(A As ListObject)
NRowOfLo = NRoZZRg(A.DataBodyRange)
End Function

Function Lon$(T)
Lon = "T_" & RmvFstNonLetter(T)
End Function

Function LoAllCol(A As ListObject) As Range
Set LoAllCol = RgzLoCC(A, 1, LoNCol(A))
End Function

Function LoAllEntCol(A As ListObject) As Range
Set LoAllEntCol = LoAllCol(A).EntireColumn
End Function

Function RgzLc(A As ListObject, C, Optional InlTot As Boolean, Optional InlHdr As Boolean) As Range
Dim R As Range
Set R = A.ListColumns(C).DataBodyRange
If Not InlTot And Not InlHdr Then
    Set RgzLc = R
    Exit Function
End If
If InlTot Then Set RgzLc = RgMoreBelow(R, 1)
If InlHdr Then Set RgzLc = RgMoreTop(R, 1)
End Function

Function DrszLoFny(L As ListObject, Fny$()) As Drs
Dim Colnoy%(): Colnoy = Cxy(FnyzLo(L), Fny)
DrszLoFny = Drs(Fny, DyzSqColnoy(SqzLo(L), Colnoy))
End Function

Function DrszLo(A As ListObject) As Drs
DrszLo = Drs(FnyzLo(A), DyzLo(A))
End Function
Function DyzLo(A As ListObject) As Variant()
DyzLo = DyoSq(SqzLo(A))
End Function

Function DyzRgColnoy(Rg As Range, Colnoy%()) As Variant()
DyzRgColnoy = DyzSqColnoy(SqzRg(Rg), Colnoy)
End Function
Function DyzLoCC(Lo As ListObject, CC) As Variant() _
' Return as many column as columns in [CC] from Lo
DyzLoCC = DyzRgColnoy(Lo.DataBodyRange, Colnoy(FnyzLo(Lo), CC))
End Function

Function DtaAdrzLo$(A As ListObject)
DtaAdrzLo = WsAdrzRg(A.DataBodyRange)
End Function

Function EntColzLo(Lo As ListObject, C) As Range ' entire col range
Set EntColzLo = RgzLc(Lo, C).EntireColumn
End Function

Function RgzLoCC(A As ListObject, C1, C2, Optional InlTot As Boolean, Optional InlHdr As Boolean) As Range
Dim R1&, R2&, mC1%, MC2%
R1 = R1Lo(A, InlHdr)
R2 = R2Lo(A, InlTot)
mC1 = WsCnozLc(A, C1)
MC2 = WsCnozLc(A, C2)
Set RgzLoCC = WsRCRC(WszLo(A), R1, mC1, R2, MC2)
End Function

Function LozWsDta(A As Worksheet, Optional Lon$) As ListObject
Set LozWsDta = NwLo(DtaRg(A), Lon)
End Function

Function FbtStrzLo$(A As ListObject)
FbtStrzLo = FbtStrzQt(A.QueryTable)
End Function

Function FnyzLo(A As ListObject) As String()
FnyzLo = Itn(A.ListColumns)
End Function

Function LoFF$(A As ListObject)
LoFF = Termln(FnyzLo(A.ListColumns))
End Function

Function HasLoC(Lo As ListObject, ColNm$) As Boolean
HasLoC = HasItn(Lo.ListColumns, ColNm)
End Function

Function IsLozNoDta(A As ListObject) As Boolean
IsLozNoDta = IsNothing(A.DataBodyRange)
End Function

Function HdrCellzLo(A As ListObject, Fldn) As Range
Dim Rg As Range: Set Rg = A.ListColumns(Fldn).Range
Set HdrCellzLo = RgRC(Rg, 1, 1)
End Function

Function LoNCol%(A As ListObject)
LoNCol = A.ListColumns.Count
End Function

Function LczLoCno(L As ListObject, C) As ListColumn
Const CSub$ = CMod & "LczLoCno"
Dim C1&: C1 = FstLc(L).DataBodyRange.Column
Dim C2&: C2 = LasLc(L).DataBodyRange.Column
If Not IsBet(C, C1, C2) Then Thw CSub, "Given-Cno is not between the FstCno & LasCno of given Lo", "Given-Cno Fst-Lo-Cno Las-Lo-Cno", C, C1, C2
Set LczLoCno = L.ListColumns(C - C1 + 1)
End Function

Function FstLc(L As ListObject) As ListColumn
Set FstLc = L.ListColumns(1)
End Function

Function LasLc(L As ListObject) As ListColumn
Set LasLc = L.ListColumns(L.ListColumns.Count)
End Function

Function Wbn$(A As Workbook)
On Error GoTo X
Wbn = A.FullName
Exit Function
X: Wbn = "WbnErr:[" & Err.Description & "]"
End Function

Function LasCWb() As Workbook
Set LasCWb = LasWb(Xls)
End Function

Function LasWb(A As Excel.Application) As Workbook
Set LasWb = A.Workbooks(A.Workbooks.Count)
End Function

Function PtzRg(A As Range, Optional Wsn$, Optional PtNm$) As PivotTable
Dim Wb As Workbook: Set Wb = WbzRg(A)
Dim Ws As Worksheet: Set Ws = AddWs(Wb)
Dim A1 As Range: Set A1 = A1zWs(Ws)
Dim Pc As PivotCache: Set Pc = WbzRg(A).PivotCaches.Create(xlDatabase, A.Address, Version:=6)
Dim Pt As PivotTable: Set Pt = Pc.CreatePivotTable(A1, DefaultVersion:=6)
End Function
Function PivCol(Pt As PivotTable, PivColNm) As PivotField

End Function
Function PivRow(Pt As PivotTable, PivRowNm) As PivotField
Set PivRow = Pt.ColumnFields(PivRowNm)
End Function
Function PivFld(A As PivotTable, F) As PivotField
Set PivFld = A.PageFields(F)
End Function
Function EntColzPt(A As PivotTable, PivColNm) As Range
Set EntColzPt = RgR(PivCol(A, PivColNm).DataRange, 1).EntireColumn
End Function
Function PivColEnt(Pt As PivotTable, ColNm) As Range
Set PivColEnt = PivCol(Pt, ColNm).EntireColumn
End Function


Function PtzLo(A As ListObject, At As Range, Rowss$, Dtass$, Optional Colss$, Optional Pagss$) As PivotTable
If WbzLo(A).FullName <> WbzRg(At).FullName Then Stop: Exit Function
Dim O As PivotTable
Set O = LoPc(A).CreatePivotTable(TableDestination:=At, TableName:=PtNmzLo(A))
With O
    .ShowDrillIndicators = False
    .InGridDropZones = False
    .RowAxisLayout xlTabularRow
End With
O.NullString = ""
SetPtffOri O, Rowss, xlRowField
SetPtffOri O, Colss, xlColumnField
SetPtffOri O, Pagss, xlPageField
SetPtffOri O, Dtass, xlDataField
Set PtzLo = O
End Function

Function PtNmzLo$(A As ListObject)

End Function

Function WbzPt(A As PivotTable) As Workbook
Set WbzPt = WbzWs(WszPt(A))
End Function

Function WszPt(A As PivotTable) As Worksheet
Set WszPt = A.Parent
End Function

Function FbtStrzQt$(A As QueryTable)
If IsNothing(A) Then Exit Function
Dim Ty As XlCmdType, Tbl$, CnStr$
With A
    Ty = .CommandType
    If Ty <> xlCmdTable Then Exit Function
    Tbl = .CommandText
    CnStr = .Connection
End With
FbtStrzQt = FmtQQ("[?].[?]", DtaSrczScvl(CnStr), Tbl)
End Function

Function CnozBefFstHid%(Ws As Worksheet)
Dim J%, O%
For J% = 1 To MaxCno
    If WsC(Ws, J).Hidden Then CnozBefFstHid = J - 1: Exit Function
Next
Stop
End Function

Function TxtCnzWc(A As WorkbookConnection) As TextConnection
On Error Resume Next
Set TxtCnzWc = A.TextConnection
End Function

Function DrzAt(At As Range) As Variant()
DrzAt = DrzSq(SqzRg(BarHzAt(At)))
End Function

Function ColzAt(At As Range) As Variant()
ColzAt = ColzSq(SqzRg(BarVzAt(At)))
End Function



Function VbarRgAt(At As Range, Optional AtLeastOneCell As Boolean) As Range
If IsEmpty(At.Value) Then
    If AtLeastOneCell Then
        Set VbarRgAt = A1zRg(At)
    End If
    Exit Function
End If
Dim R1&: R1 = At.Row
Dim R2&
    If IsEmpty(RgRC(At, 2, 1)) Then
        R2 = At.Row
    Else
        R2 = At.End(xlDown).Row
    End If
Dim C%: C = At.Column
Set VbarRgAt = WsCRR(WszRg(At), C, R1, R2)
End Function

Function FstWsn$(Fx)
FstWsn = FstItm(WnyzFx(Fx))
End Function

Function OleCnStrzFx$(Fx)
OleCnStrzFx = "OLEDb;" & AdoCnStrzFx(Fx)
End Function

Function HasFx(Fx) As Boolean
Dim Wb As Workbook
For Each Wb In Xls.Workbooks
    If Wb.FullName = Fx Then HasFx = True: Exit Function
Next
End Function

Private Sub RgMoreBelow__Tst()
Dim R As Range, Act As Range, Ws As Worksheet
Set Ws = NwWs
Set R = Ws.Range("A3:B5")
Set Act = RgMoreTop(R, 1)
Debug.Print Act.Address
Stop
Debug.Print RgRR(R, 1, 2).Address
Stop
End Sub

Function AutoFilterzLo(L As ListObject) As AutoFilter
Dim A: A = L.AutoFilter
If IsNothing(A) Then Stop
Set AutoFilterzLo = A
End Function

Function CvAutoFilter(A) As AutoFilter
Set CvAutoFilter = A
End Function

Function BarHzAt(At As Range) As Range
Dim A1 As Range: Set A1 = A1zRg(At)
If IsEmpty(A1.Value) Then Set BarHzAt = A1: Exit Function
Dim C2&: C2 = A1.End(xlRight).Column - A1.Column + 1
Set BarHzAt = RgCRR(A1, 1, 1, C2)
End Function

Function BarVzAt(At As Range) As Range
Dim A1 As Range: Set A1 = A1zRg(At)
If IsEmpty(A1.Value) Then Set BarVzAt = A1: Exit Function
Dim R2&: R2 = A1.End(xlDown).Row - A1.Row + 1
Set BarVzAt = RgCRR(A1, 1, 1, R2)
End Function

Function A1(Ws As Worksheet) As Range
Set A1 = WsRC(Ws, 1, 1)
End Function

Function ColzRg(Rg, C) As Variant()
Dim R As Range
ColzRg = ColzSq(SqzRg(R))
End Function

Sub SwapCellVal(Cell1 As Range, Cell2 As Range)
Dim A: A = RgRC(Cell1, 1, 1).Value
RgRC(Cell1, 1, 1).Value = RgRC(Cell2, 1, 1).Value
RgRC(Cell2, 1, 1).Value = A
End Sub


Function ResiRg(At As Range, Sq()) As Range
If Si(Sq) = 0 Then Set ResiRg = A1zRg(At): Exit Function
Set ResiRg = RgRCRC(At, 1, 1, NRowOfSq(Sq), NColzSq(Sq))
End Function

Function EmpRxy(Sq()) As Long() '#Empty-Rxy#
Dim Lc%: Lc = LBound(Sq, 2)
Dim Uc%: Uc = UBound(Sq, 2)
Dim R&: For R = LBound(Sq, 1) To UBound(Sq, 1)
    If IsEmpRow(Sq, R, Uc, Lc) Then PushI EmpRxy, R
Next
End Function

Function IsEmpRow(Sq(), R&, Lc%, Uc%) As Boolean
Dim C%: For C = Lc To Uc
    If Not IsEmpty(Sq(R, C)) Then Exit Function
Next
IsEmpRow = True
End Function

Function SqzRgNo(A As Range) As Variant()
'Ret : #Sq-Fm:Rg-How:No ! a sq fm:@A how:no means the @ret:sq is using Rno & Cno as index.
Dim O(): O = SqzRg(A)
Dim R1&, R2&, C1%, C2%
ReDim Preserve O(R1 To R2, C1 To C2)
SqzRgNo = O
End Function

Function WbzRg(A As Range) As Workbook
Set WbzRg = WbzWs(WszRg(A))
End Function

Function WszRg(A As Range) As Worksheet
Set WszRg = A.Parent
End Function


Function DrszLon(Ws As Worksheet, Lon$) As Drs
DrszLon = DrszLo(Ws.ListObjects(Lon))
End Function

Function ColzLc(Lc As ListColumn) As Variant()
ColzLc = ColzSq(SqzRg(Lc.DataBodyRange))
End Function

Function ColzLo(Lo As ListObject, C) As Variant()
ColzLo = ColzLc(Lo.ListColumns(C))
End Function

Function StrColzLo(Lo As ListObject, C) As String()
StrColzLo = StrColzLc(Lo.ListColumns(C))
End Function

Function StrColzLc(Lc As ListColumn) As String()
StrColzLc = StrColzSq(SqzRg(Lc.DataBodyRange))
End Function

Function StrColzWsLC(Ws As Worksheet, Lon$, C$) As String()
StrColzWsLC = StrColzLo(Ws.ListObjects(Lon), C)
End Function

Function VbarIntAy(A As Range) As Integer()
'VbarIntAy = AyIntAy(VbarAy(A))
End Function

Function TmpDbzFx(Fx) As Database
Set TmpDbzFx = TmpDbzFxWny(Fx, WnyzFx(Fx))
End Function

Function TmpDbzFxWny(Fx, Wny$()) As Database
Dim O As Database
   Set O = TmpDb
Dim W: For Each W In Itr(Wny)
    LnkFxw O, ">" & W, Fx, W
Next
Set TmpDbzFxWny = O
End Function

Function HasWb(Wbn) As Boolean
Dim B As Workbook: For Each B In Xls.Workbooks
    If B.Name = Wbn Then HasWb = True: Exit Function
Next
End Function

Function NoWb(Wbn) As Boolean
Const CSub$ = CMod & "NoWb"
NoWb = Not HasWb(Wbn)
If NoWb Then InfLn CSub, FmtQQ("Wbn[?] not found", Wbn)
End Function

Function Wb(Wbn) As Workbook
Set Wb = Xls.Workbooks(Wbn)
End Function

Function WbzFx(Fx) As Workbook
Set WbzFx = Wb(Fx)
End Function

Function WszFxw(Fx, Optional Wsn$ = "Data") As Worksheet
Set WszFxw = WszWb(WbzFx(Fx), Wsn)
End Function

Function ArszFx(Fx$, W$, C) As ADODB.Recordset
Set ArszFx = ArszCnq(CnzFx(Fx), SqlSel_F_T(C, AxTbn(W)))
End Function

Function ArszFxDist(Fx$, W$, DistC$) As ADODB.Recordset
Set ArszFxDist = ArszCnq(CnzFx(Fx), SqlSel_FF_Fm(DistC, AxTbn(W), Dis:=True))
End Function

Function WsCdNyzFx(Fx) As String()
Dim Wb As Workbook
Set Wb = WbzFx(Fx)
WsCdNyzFx = WsCdNy(Wb)
Wb.Close False
End Function

Private Sub FstWsn__Tst()
Dim Fx$
Fx = SalTxtFx
Ept = "Sheet1"
GoSub T1
Exit Sub
T1:
    Act = FstWsn(Fx)
    C
    Return
End Sub

Private Sub TmpDbzFx__Tst()
Dim Db As Database: Set Db = TmpDbzFx(MB52LasIFx)
DmpAy Tny(Db)
Db.Close
End Sub

Function FxwwMisEr(Fx, WW, Optional FxKd$ = "Excel file") As String()
':WW: :TermLn ! #Worksheet-Name-TermLn#
FxwwMisEr = EoFfnMis(Fx, FxKd)
Dim W: For Each W In Termy(WW)
    PushIAy FxwwMisEr, FxwMisEr(Fx, W, FxKd)
Next
End Function

Function FxwMisEr(Fx, W, Optional FilKd$ = "Excel file") As String()
If HasFxw(Fx, W) Then FxwMisEr = FxwMisMsg(Fx, W, FilKd)
End Function

Function FxwMisMsg(Fx, W, Optional FilKd$ = "Excel file") As String()
BfrV FmtQQ("[?] miss ws [?]", FilKd, W)
BfrTab "Path  : " & Pth(Fx)
BfrTab "File  : " & Fn(Fx)
BfrTab "Has Ws: " & Termln(Wny(Fx))
FxwMisMsg = BfrLy
End Function

Function WszAy2(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2") As Worksheet
Set WszAy2 = WszDrs(DrsFmAy2(A, B, N1, N2))
End Function

Function WszCd(Wb As Workbook, WsCdn$) As Worksheet
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    If Ws.CodeName = WsCdn Then Set WszCd = Ws: Exit Function
Next
End Function
Function NyzFml(Fml$) As String()
NyzFml = MacroNy(Fml)
End Function

Function WszLy(Ly$(), Optional Wsn$ = "Sheet1") As Worksheet
Set WszLy = WszRg(RgzAyV(Ly, A1zWs(NwWs(Wsn))))
End Function

Function SqHzN(N%) As Variant()
Dim O()
ReDim O(1 To 1, 1 To N)
SqHzN = O
End Function

Function SqVzN(N%) As Variant()
Dim O(), J%
ReDim O(1 To N, 1 To 1)
SqVzN = O
End Function

Function N_ZerFill$(N, NDig&)
N_ZerFill = Format(N, String(NDig, "0"))
End Function

Function WszS12y(A() As S12, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Worksheet
Set WszS12y = WszSq(SqzS12y(A, Nm1, Nm2))
End Function

Private Sub Ay2Ws__Tst()
GoTo Z
Dim A, B
Z:
    A = SyzSS("A B C D E")
    B = SyzSS("1 2 3 4 5")
    VisWs WszAy2(A, B)
End Sub

Private Sub WbFbOupTbl__Tst()
GoTo Z
Z:
    Dim W As Workbook
    'Set W = WbFbOupTbl(WFb)
    'VisWb W
    Stop
    'W.Close False
    Set W = Nothing
End Sub

Function LasCnozRg%(R As Range)
LasCnozRg = R.Column + NColzRg(R) - 1
End Function

Function LasRnozRg&(R As Range)
LasRnozRg = R.Row + NRoZZRg(R) - 1
End Function

Function FnyzWs(A As Worksheet) As String()
FnyzWs = FnyzLo(FstLo(A))
End Function

Function HasWbn(Xls As Excel.Application, Wbn$) As Boolean
HasWbn = HasItn(Xls.Workbooks, Wbn)
End Function

Function ColRgAy(L As ListObject, Colnoy$()) As Range()
'Ret : :ColAy:RgAy: of each @L-col stated in @Colnoy.
Dim C: For Each C In Itr(Colnoy)
    PushObj ColRgAy, L.ListColumns(C).DataBodyRange
Next
End Function

Function ColAyzLo(Lo As ListObject, Cxy) As Variant()
'Fm Cxy : #Col-iX-aY ! a col-ix can be a number running fm 1 or a coln.
'Ret    : #Col-Ay    ! ay-of-col.  A col is ay-of-val-of-a-col.  All col has same # of ele. @@
Dim C: For Each C In Itr(Cxy)
    Dim Lc As ListColumn: Set Lc = Lo.ListColumns(C)
    PushI ColAyzLo, ColzLo(Lo, C)
Next
End Function

Sub AddHypLnk(Rg As Range, Wsn)
Dim A1 As Range: Set A1 = WszWb(WbzRg(Rg), Wsn).Range("A1")
Rg.HypErLnks.Add Rg, "", SubAddress:=A1.Address(External:=True)
End Sub

Function FilterzLo(Lo As ListObject, Coln)
'Ret : Set filter of all Lo of CWs @
Dim Ws  As Worksheet:   Set Ws = CWs
Dim C$:                      C = "Mthn"
Dim Lc  As ListColumn:  Set Lc = Lo.ListColumns(C)
Dim OFld%:                OFld = Lc.Index
Dim Itm():                 Itm = ColzLc(Lc)
Dim Patn$:                Patn = "^Ay"
Dim OSel:                 OSel = AwPatn(Itm, Patn)
Dim ORg As Range:      Set ORg = Lo.Range
ORg.AutoFilter Field:=OFld, Criteria1:=OSel, Operator:=xlFilterValues
End Function

Function LozRg(R As Range) As ListObject
Dim R1 As Range: Set R1 = RgRC(R, 2, 1)
Dim L As ListObject: For Each L In WszRg(R).ListObjects
    If HasRg(L, R1) Then Set LozRg = L: Exit Function
Next
End Function
