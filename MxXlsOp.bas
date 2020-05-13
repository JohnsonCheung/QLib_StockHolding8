Attribute VB_Name = "MxXlsOp"
Option Compare Text
Option Explicit
Const CNs$ = "Xls.Op"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsOp."
Enum eWsPos: eWsAtEnd: eWsAtBeg: eWsBef: eWsAft: End Enum

Sub BrwFx(Fx)
ChkFfnExist Fx
VisWb WbzFx(Fx)
End Sub

Sub CrtFx(Fx): WbSavAs(NwWb, Fx).Close: End Sub

Sub EnsFx(Fx) ' Crt an emp @Fx if not exist.  Ret @Fx
If NoFfn(Fx) Then CrtFx Fx
End Sub

Sub InsRow(At As Range, Optional N = 1)
EntRgRR(At, 1, 1 + N - 1).EntireRow.Insert ' @OCell will be changed after insert  ! <==
End Sub

Sub RmvEmpRow(R As Range)
Dim Ws As Worksheet: Set Ws = WszRg(R)
Dim Sq():                Sq = SqzRgNo(R)
Dim Rxy&():             Rxy = EmpRxy(Sq)
EntRows(Ws, Rxy).Remove
End Sub

Private Function EntRowsStr$(Rny&())
Dim O$()
Dim Rno: For Each Rno In Rny
    PushI O, Rno & ":" & Rno
Next
EntRowsStr = JnComma(O)
End Function

Function EntRows(Ws As Worksheet, Rny&()) As Range: Set EntRows = Ws.Range(EntRowsStr(Rny)): End Function
Sub RmvNRow(At As Range, Optional N = 1): EntRgRR(At, 1, N).Delete: End Sub

Sub OpnFxy(Fxy$()): OpnXFxy NwXls, Fxy: End Sub
Sub OpnXFxy(X As Excel.Application, Fxy$())
MinvXls X
Dim F: For Each F In Fxy
    X.Workbooks.Open F
Next
VArrangeWb X
End Sub

Function OpnFxMax(Fx) As Workbook
Set OpnFxMax = OpnFx(Fx)
MaxvWb OpnFxMax
End Function

Function OpnFx(Fx) As Workbook:  Set OpnFx = OpnXFx(NwXlsMinv, Fx): End Function

Function OpnFxIfExist(Fx) As Boolean
If HasFfn(Fx) Then OpnFx Fx: OpnFxIfExist = True
End Function

Function OpnXFx(X As Excel.Application, Fx) As Workbook
ChkFfnExist Fx, CSub
Set OpnXFx = X.Workbooks.Open(Fx)
End Function

Sub ClrWsNm(Ws As Worksheet)
Dim N: For Each N In Itn(Ws.Names)
    Ws.Names(N).Delete
Next
End Sub
'---
Private Sub EnsHyp__Tst()
Dim Rg As Range
GoSub Crt
GoSub T0
GoSub Rmv
Exit Sub
T0:
    Set Rg = CWs.Range("A1:A2")
    EnsHyp Rg
    'Stop
    Return
Crt:
    Dim Wb As Workbook: Set Wb = NwWb("A")
    AddWs Wb, "A 1"
    AddWs Wb, "A 2"
    WszWb(Wb, "A").Activate
    CWs.Range("A1").Value = "A 1"
    CWs.Range("A2").Value = "A 2"
    VisWb Wb
    Return
Rmv:
    Wb.Close False
    Return
End Sub

Sub EnsHyp(Rg As Range) ' if cell of @Rg has value of cur wb wsn, ens that cell hyp lnk pointing the ws A1
Dim W$(): W = Wny(WbzRg(Rg))
Dim Cell As Range: For Each Cell In Rg
    W1Ens Cell, W
Next
End Sub

Private Sub W1Ens(Cell As Range, W$()) ' If the @Cell.Value is in @W, ens @Cell's hyp lnk to point to that ws A1
Dim Wsn$
    Dim V:  V = Cell.Value
    If IsStr(V) Then If Not HasEle(W, V) Then Exit Sub
    Wsn = V  ' The wsn pointed by the @Cell.Value, if the @Cell.Value is str and in @W
Dim Adr$                                                   ' The adr of A1 of %Wsn
    If Wsn <> "" Then Adr = WszWb(WbzRg(Cell), Wsn).Range("A1").Address(External:=True)

Dim Has As Boolean
    Dim L As Hyperlink: For Each L In Cell.Hyperlinks
        If L.SubAddress = Adr Then
            Has = True
        Else
            Cell.HypErLnks.Delete
        End If
    Next
If Has Then Exit Sub
Cell.HypErLnks.Add Anchor:=Cell, Address:="", SubAddress:=Adr '<== Add Hyp lnk
End Sub
'------
Sub RmvWsIf(Fx, Wsn$)
'Ret : Ret @Wsn in @Fx if exists @@
If HasFxw(Fx, Wsn) Then
   Dim B As Workbook: Set B = WbzFx(Fx)
   WszWb(B, Wsn).Delete
   SavWb B
   ClsCWbNoSav B
End If
End Sub

Function LozAyH(Ay, Wb As Workbook, Optional Wsn$, Optional Lon$) As ListObject
Set LozAyH = NwLo(RgzSq(Sqh(Ay), A1zWb(Wb, Wsn)), Lon)
End Function

Private Sub SetWscdn__Tst()
Dim A As Worksheet: Set A = NwWs
SetWscdn A, "XX"
MaxiWs A
End Sub

Sub MgeBottomCell(VBar As Range)
Ass IsSngCol(VBar)
Dim R2: R2 = NRoZZRg(VBar)
Dim R1
    Dim Fnd As Boolean
    For R1 = R2 To 1 Step -1
        If Not IsEmpty(RgRC(VBar, R1, 1)) Then Fnd = True: GoTo Nxt
    Next
Nxt:
    If Not Fnd Then Stop
If R2 = R1 Then Exit Sub
Dim R As Range: Set R = RgCRR(VBar, 1, R1, R2)
R.Merge
R.VerticalAliment = XlVAlign.xlVAlignTop
End Sub

Sub SetWscdn(A As Worksheet, CdNm$)
CmpzWs(A).Name = CdNm
End Sub

Sub SetWsCdnAndLon(A As Worksheet, Nm$)
CmpzWs(A).Name = Nm
SetLon FstLo(A), Nm
End Sub

Function HasCell(A As Range, Cell As Range) As Boolean
If Not IsCell(Cell) Then Exit Function
If Not IsBet(Cell.Row, A.Row, NRoZZRg(A)) Then Exit Function
If Not IsBet(Cell.Column, A.Column, NColzRg(A)) Then Exit Function
HasCell = True
End Function

Sub MgeRg(A As Range)
A.MergeCells = True
A.HorizontalAliment = XlHAlign.xlHAlignCenter
A.VerticalAliment = Excel.XlVAlign.xlVAlignCenter
End Sub

Sub ClsAllCWbNoSav()
Dim X As Excel.Application: Set X = Xls
While X.Workbooks.Count > 0
    ClsCWbNoSav X.Workbooks(1)
Wend
End Sub
Sub ClsLasCWbNoSav()
ClsCWbNoSav LasCWb
End Sub
Sub ClsCWbNoSav(A As Workbook)
A.Close False
End Sub

Sub ClsCWsNoSav(A As Worksheet)
WbzWs(A).Close False
End Sub

Sub MaxvWb(A As Workbook): VisWb A: MaxiWb A: End Sub

Sub MaxiWb(A As Workbook): A.Application.WindowState = xlMaximized: End Sub

Function MaxiWs(A As Worksheet) As Worksheet
A.Application.WindowState = xlMaximized
Set MaxiWs = A
End Function

Sub QuitWb(A As Workbook): QuitXls A.Application: End Sub

Function SavWbCsv(A As Workbook, Fcsv$) As Workbook
DltFfnIf Fcsv
A.Application.DisplayAlerts = False
A.SaveAs Fcsv, XlFileFormat.xlCSV
A.Application.DisplayAlerts = True
Set SavWbCsv = A
End Function

Function SavWb(A As Workbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.Save
A.Application.DisplayAlerts = Y
Set SavWb = A
End Function

Function WbSavAs(A As Workbook, Fx, Optional Fmt As XlFileFormat = xlOpenXMLWorkbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.SaveAs Fx, Fmt
A.Application.DisplayAlerts = Y
Set WbSavAs = A
End Function

Sub SetWcFcsv(A As Workbook, Fcsv$)
'Set first Wb TextConnection to Fcsv if any
Dim T As TextConnection: Set T = TxtWc(A)
Dim C$: C = T.Connection: If Not HasPfx(C, "TEXT;") Then Stop
T.Connection = "TEXT;" & Fcsv
End Sub

Function HasWsCd(WsCdn$) As Boolean
HasWsCd = HasItp(CWb.Sheets, "CodeName", WsCdn)
End Function
Function HasWs(A As Workbook, WsIx) As Boolean
If IsNumeric(WsIx) Then
    HasWs = IsBet(WsIx, 1, A.Sheets.Count)
    Exit Function
End If
Dim Ws As Worksheet
For Each Ws In A.Worksheets
    If Ws.Name = WsIx Then HasWs = True: Exit Function
Next
End Function

Private Sub WbWcsy__Tst()
'D WcStrAyWbOLE(WbzFx(TpFx))
End Sub

Private Sub LozAyH__Tst()
'D NyOy(LozAyH(TpWb))
End Sub

Private Sub TxtWcCnt__Tst()
Dim O As Workbook: 'Set O = WbzFx(Vbe_MthFx)
Ass TxtWcCnt(O) = 1
O.Close
End Sub

Private Sub SetWcFcsv__Tst()
Dim Wb As Workbook
'Set Wb = WbzFx(Vbe_MthFx)
Debug.Print TxtWcStr(Wb)
SetWcFcsv Wb, "C:\ABC.CSV"
Ass TxtWcStr(Wb) = "TEXT;C:\ABC.CSV"
Wb.Close False
Stop
End Sub


Private Sub Clr_LoRow__Tst()
DltLoRow CWs.ListObjects("T_SrcCd")
End Sub
Sub DltLoRow(A As ListObject)
Dim R As Range: Set R = A.DataBodyRange
If IsNothing(R) Then Exit Sub
R.ClearContents
Set R = A1zRg(A.ListColumns(1).Range)
Dim R1 As Range: Set R1 = RgRR(R, 1, 2)
A.Resize R1
End Sub

Sub DltLo(A As Worksheet)
Dim Ay() As ListObject, J%
Ay = IntozItr(Ay, A.ListObjects)
For J = 0 To UB(Ay)
    Ay(J).Delete
Next
End Sub

Sub DltWsIf(A As Workbook, WsIx)
If HasWs(A, WsIx) Then DltWs A, WsIx
End Sub

Sub SavAsCls(Wb As Workbook, Fx)
Wb.SaveAs Fx
Wb.Close
End Sub

Sub SavAsFxm(Wb As Workbook, Fxm)
Wb.SaveAs Fxm, XlFileFormat.xlOpenXMLWorkbookMacroEnabled
End Sub

Function SavAsTmpFxm$(Wb As Workbook)
Dim O$: O = TmpFxm
SavAsFxm Wb, O
SavAsTmpFxm = O
End Function

Function WbnzWs$(A As Worksheet)
WbnzWs = WbzWs(A).FullName
End Function

Sub DltColFm(Ws As Worksheet, FmCol)
WsCC(Ws, FmCol, LasCno(Ws)).Delete
End Sub
Sub DltRowFm(Ws As Worksheet, FmRow)
WsRR(Ws, FmRow, LasRno(Ws)).Delete
End Sub
Sub HidColFm(Ws As Worksheet, FmCol)
WsCC(Ws, FmCol, MaxCno).Hidden = True
End Sub

Sub HidRowFm(Ws As Worksheet, FmRow&)
WsRR(Ws, FmRow, MaxRno).EntireRow.Hidden = True
End Sub


Function PtCpyToLo(A As PivotTable, At As Range) As ListObject
Dim R1, R2, C1, C2, NC, NR
    R1 = A.RowRange.Row
    C1 = A.RowRange.Column
    R2 = LasRnozRg(A.DataBodyRange)
    C2 = LasCnozRg(A.DataBodyRange)
    NC = C2 - C1 + 1
    NR = R2 - C1 + 1
WsRCRC(WszPt(A), R1, C1, R2, C2).Copy
At.PasteSpecial xlPasteValues

Set PtCpyToLo = NwLo(RgRCRC(At, 1, 1, NR, NC))
End Function

Sub SetPtffOri(A As PivotTable, FF$, Ori As XlPivotFieldOrientation)
Dim F, J%, T
T = Array(False, False, False, False, False, False, False, False, False, False, False, False)
J = 1
For Each F In Itr(FnyzFF(FF))
    With PivFld(A, F)
        .Orientation = Ori
        .Position = J
        If Ori = xlColumnField Or Ori = xlRowField Then
            .Subtotals = T
        End If
    End With
    J = J + 1
Next
End Sub

Sub ChkWsnExist(Wb As Workbook, Wsn$, Fun$)
If HasWs(Wb, Wsn) Then
    Thw Fun, "Wb should have not have Ws", "Wb Ws", Wb.FullName, Wsn
End If
End Sub

Sub SetPtWdt(A As PivotTable, Colss$, ColWdt As Byte)
If ColWdt <= 1 Then Stop
Dim C
For Each C In Itr(SyzSS(Colss))
    EntColzPt(A, C).ColumnWidth = ColWdt
Next
End Sub

Sub SetPtOutLin(A As PivotTable, Colss$, Optional Lvl As Byte = 2)
If Lvl <= 1 Then Stop
Dim F, C As VBComponent
For Each C In Itr(SyzSS(Colss))
    EntColzPt(A, F).OutlineLevel = Lvl
Next
End Sub

Sub SetPtRepeatLbl(A As PivotTable, Rowss$)
Dim F
For Each F In Itr(SyzSS(Rowss))
    PivFld(A, F).RepeatLabels = True
Next
End Sub

Sub ShwPt(A As PivotTable)
VisXls A.Application
End Sub

Function NwA1(Optional Wsn$) As Range
Set NwA1 = A1zWs(NwWs(Wsn))
End Function

Function NwWb(Optional Wsn$) As Workbook
Dim O As Workbook
Set O = NwXls.Workbooks.Add
Set NwWb = WbzWs(SetWsn(FstWs(O), Wsn))
End Function

Function NwWszAy(Ay, Optional Hdr$ = "An-Array", Optional Wsn$) As Worksheet
Dim O As Worksheet: Set O = NwWs(Wsn)
PutAyv AddItmAy(Hdr, Ay), A1zWs(O)
Set NwWszAy = O
End Function

Function NwWs(Optional Wsn$) As Worksheet
Set NwWs = SetWsn(FstWs(NwWb), Wsn)
End Function

Function XlszGet() As Excel.Application
Set XlszGet = GetObject(, "Excel.Application")
End Function

Function NwXlszFx(Fx) As Excel.Application
Dim O As Excel.Application: Set O = NwXls
O.Workbooks.Open Fx
Set NwXlszFx = O
End Function

Function NwWbzFx(Fx$) As Workbook
Set NwWbzFx = NwXls.Workbooks.Open(Fx, UpdateLinks:=False)
NwWbzFx.Application.DisplayAlerts = False
End Function

Function NwXls() As Excel.Application ' CrtObj Xls.App
Dim O As Excel.Application
Set O = CreateObject("Excel.Application") ' Don't use New Excel.Application, but why?
Set NwXls = O
End Function

Function NwXlsMinv() As Excel.Application ' CrtObj Xls.App & Minv
Set NwXlsMinv = MinvXls(NwXls)
End Function

Private Sub NwXls__Tst()
NwXls
Stop
Exit Sub
Dim A
'{00024500-0000-0000-C000-000000000046}
Set A = Interaction.CreateObject("{00024500-0000-0000-C000-000000000046}", "Excel.Application")
Stop
End Sub

Sub QuitXls(A As Excel.Application)
Stamp "QuitXls: Start"
Stamp "QuitXls: ClsAllWb":    ClsAllWb A
Stamp "QuitXls: Quit":        A.Quit
Stamp "QuitXls: Set nothing": Set A = Nothing
Stamp "QuitXls: Done"
End Sub

Sub ClsAllWb(A As Excel.Application)
Dim W As Workbook
For Each W In A.Workbooks
    W.Close False
Next
End Sub


Sub RplLozFb(Wb As Workbook, Fb)
Dim Ws As Worksheet, D As Database
Set D = Db(Fb)
For Each Ws In Wb.Sheets
    RplLozWs Ws, D
Next
D.Close
End Sub

Sub RplLozWs(Ws As Worksheet, Optional D As Database)
Dim Lo As ListObject
For Each Lo In Ws.ListObjects
    RplLozT Lo, "@" & Mid(Lo.Name, 3), D
Next
End Sub

Sub RplLozT(A As ListObject, T, Optional D As Database)
Const CSub$ = CMod & "RplLozT"
Dim Fny1$(): Fny1 = Fny(D, T)
Dim Fny2$(): Fny2 = FnyzLo(A)
If Not IsAySam(Fny1, Fny2) Then
    Thw CSub, "LoFny and TblFny are not same", "LoFny T TblFny Db", Fny2, T, Fny1, D.Name
End If
Dim Sq()
    Dim R As DAO.Recordset
    Set R = Rs(A, SqlSel_Fny_T(Fny2, T))
    Sq = AddSngQuozSq(SqzRs(R))
MinxLo A
RgzSq Sq, A.DataBodyRange
End Sub

Private Sub NwFxzOupTbl__Tst()
Dim Fx$: Fx = TmpFx
NwFxzOupTbl Fx, DutyDtaFb
OpnFx Fx
End Sub

Sub NwFxzOupTbl(Fx, Fb, Optional Way As eWsdAddWay)
SavAsCls NwWbzFbOup(Fb, Way), Fx
End Sub

Sub PutSnoDown(At As Range, N, Optional Fm = 1)
PutAyv LngSno(N - 1, Fm), At
End Sub

Sub DltSheet1(Wb As Workbook)
DltWs Wb, "Sheet1"
End Sub
Sub ActWs(Ws As Worksheet)
If IsEqObj(Ws, CWs) Then Exit Sub
Ws.Activate
End Sub
Sub DltWs(Wb As Workbook, WsIx)
Wb.Application.DisplayAlerts = False
If Wb.Sheets.Count = 1 Then Exit Sub
If HasWs(Wb, WsIx) Then WszWb(Wb, WsIx).Delete
End Sub

Sub ClrDown(A As Range)
VbarRgAt(A, AtLeastOneCell:=True).Clear
End Sub


Sub MgeCellAbove(Cell As Range)
'If Not IsEmpty(A.Value) Then Exit Sub
If Cell.MergeCells Then Exit Sub
If Cell.Row = 1 Then Exit Sub
If RgRC(Cell, 0, 1).MergeCells Then Exit Sub
MgeRg RgCRR(Cell, 1, 0, 1)
End Sub


Sub FillSeqH(HBar As Range)
Dim Sq()
Sq = SqVzN(NRoZZRg(HBar))
ResiRg(HBar, Sq).Value = Sq
End Sub
Sub ClrCellBelow(Cell As Range)
CellBelow(Cell).Clear
End Sub

Sub FillSeqV(VBar As Range)
Dim Sq()
Sq = SqHzN(NRoZZRg(VBar))
ResiRg(VBar, Sq).Value = Sq
End Sub

Sub FillWny(At As Range)
RgzAyV Wny(WbzRg(At)), At
End Sub

Sub FillAtV(At As Range, Ay)
FillSq Sqv(Ay), At
End Sub

Sub FillLc(Lo As ListObject, ColNm$, Ay)
Const CSub$ = CMod & "FillLc"
If NRoZZRg(Lo.DataBodyRange) <> Si(Ay) Then Thw CSub, "Lo-NRow <> Si(Ay)", "Lo-NRow ,Si(Ay)", NRowOfLo(Lo), Si(Ay)
Dim At As Range, C As ListColumn, R As Range
'DmpAy FnyzLo(Lo)
'Stop
Set C = Lo.ListColumns(ColNm)
Set R = C.DataBodyRange
Set At = R.Cells(1, 1)
FillAtV At, Ay
End Sub
Sub FillSq(Sq(), At As Range)
ResiRg(At, Sq).Value = Sq
End Sub
Sub FillAtH(Ay, At As Range)
FillSq Sqh(Ay), At
End Sub

Sub RunFxqByCn(Fx, Q)
CnzFx(Fx).Execute Q
End Sub
Function DKVzKSet(KSet As Dictionary) As Drs
Dim K, Dy(): For Each K In KSet.Keys
    Dim Sset As Dictionary: Set Sset = KSet(K)
    Dim V: For Each V In Sset.Keys
        PushI Dy, Array(K, V)
    Next
Next
DKVzKSet = DrszFF("K V", Dy)
End Function
Private Sub DKVzLoFilter__Tst()
Dim Lo As ListObject: Set Lo = FstLo(CWs)
BrwDrs DKVzLoFilter(Lo)
End Sub
Function DKVzLoFilter(L As ListObject) As Drs
DKVzLoFilter = DKVzKSet(KSetzLoFilter(L))
End Function

Function KSetzKyAetAy(Ky$(), AetAy() As Dictionary) As Dictionary
Set KSetzKyAetAy = New Dictionary
Dim K, J&: For Each K In Itr(Ky)
    KSetzKyAetAy.Add K, AetAy(J)
    J = J + 1
Next
End Function
Sub SetOnFilter(L As ListObject)
On Error GoTo X
Dim M As Boolean: M = L.AutoFilter.FilterMode ' If filter is on, it will have no err, otherwise, there is err
Exit Sub
X:
L.Range.AutoFilter 'Turn on
End Sub
Function KSetzLoFilter(L As ListObject) As Dictionary
'Ret : KSet
Dim O As Dictionary: Set O = New Dictionary
SetOnFilter L
Dim Fny$(): Fny = FnyzLo(L)
Dim F As Filter, J%: For Each F In L.AutoFilter.Filters
    Dim K$: K = Fny(J)
    KSetzLoFilter__Add O, K, F
    J = J + 1
Next
Set KSetzLoFilter = O
End Function

Sub KSetzLoFilter__Add(OKSet As Dictionary, K$, F As Filter)
If Not F.On Then Exit Sub
If F.Operator <> xlFilterValues Then Exit Sub
Dim S As Dictionary: Set S = Aet(AmRmvPfx(F.Criteria1, "="))
OKSet.Add K, S
End Sub

Function NwWbzDrs(D As Drs) As Workbook
Set NwWbzDrs = WbzRg(RgzDrs(D, NwA1))
End Function

Function RgzDrs(A As Drs, At As Range) As Range
Set RgzDrs = RgzSq(SqzDrs(A), At)
End Function

Function LozDrs(A As Drs, At As Range, Optional Lon$) As ListObject
Set LozDrs = NwLo(RgzDrs(A, At), Lon)
End Function

Function WszAy(Ay, Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet, R As Range
Set O = NwWs(Wsn)
O.Range("A1").Value = "Array"
Set R = RgzSq(Sqv(Ay), O.Range("A2"))
NwLo RgMoreTop(R)
Set WszAy = O
End Function

Function WszDrs(A As Drs, Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NwWs(Wsn)
Dim L As ListObject: Set L = LozDrs(A, O.Range("A1"))
Dim Lc As ListColumn: For Each Lc In L.ListColumns
    StdFmtLc Lc
Next
Set WszDrs = O
End Function

Function RgzAyV(Ay, At As Range) As Range
Set RgzAyV = RgzSq(Sqv(Ay), At)
End Function

Function RgzAyH(Ay, At As Range) As Range
Set RgzAyH = RgzSq(Sqh(Ay), At)
End Function

Function RgzDy(Dy(), At As Range) As Range
Set RgzDy = RgzSq(SqzDy(Dy), At)
End Function

Function WszDy(Dy(), Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NwWs(Wsn)
RgzDy Dy, A1zWs(O)
Set WszDy = O
End Function

Function WszDs(A As Ds) As Worksheet
Dim O As Worksheet: Set O = NwWs
A1zWs(O).Value = "*Ds " & A.DsNm
Dim At As Range: Set At = WsRC(O, 2, 1)
Dim BelowN&
Dim Ay() As Dt: Ay = A.DtAy
Dim J&: For J = 0 To DtUB(Ay)
    Dim Dt As Dt: Dt = Ay(J)
    LozDt Dt, At
    BelowN = 2 + Si(Dt.Dy)
    Set At = CellBelow(At, BelowN)
Next
Set WszDs = O
End Function

Function RgzDt(A As Dt, At As Range, Optional DtIx%)
Dim Pfx$: If DtIx > 0 Then Pfx = QuoBkt(CStr(DtIx))
At.Value = Pfx & A.DtNm
RgzSq SqzDrs(DrszDt(A)), CellBelow(At)
End Function

Function LozDt(A As Dt, At As Range) As ListObject
Dim R As Range
If At.Row = 1 Then
    Set R = RgRC(At, 2, 1)
Else
    Set R = At
End If
Set LozDt = LozDrs(DrszDt(A), R)
RgRC(R, 0, 1).Value = A.DtNm
End Function


Private Sub WszDs__Tst()
VisWs WszDs(SampDs)
End Sub

Function EnsWbn(Wbn$, InXls As Excel.Application) As Workbook
Dim O As Workbook
Const FxFn$ = "Insp.xlsx"
If HasWbn(InXls, FxFn) Then
    Set O = InXls.Workbooks(FxFn)
Else
    Set O = InXls.Workbooks.Add
    O.SaveAs InstFdr("Insp") & "Insp.xlsx"
End If
Set EnsWbn = VisWb(O)
End Function

Sub OpnFcsv(Fcsv)
Xls.Workbooks.OpenText Fcsv
End Sub
