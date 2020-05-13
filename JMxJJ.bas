Attribute VB_Name = "JMxJJ"
Const CLib$ = "QAppMB52."
#If False Then
Option Compare Text
Option Explicit
Const CMod$ = CLib & "JMxJJ."
Type Sy12
    Sy1() As String
    Sy2() As String
End Type
Function NReczT&(T)
NReczT = VzQ("Select Count(*) from [" & T & "]")
End Function

Function AlignR$(S, W%)
AlignR = Space(W - Len(S)) & S
End Function

Function IsEqStr(A, B, Optional C As VbCompareMethod = vbBinaryCompare) As Boolean
IsEqStr = StrComp(A, B, C) = 0
End Function

Function Max(A, B)
If A > B Then
    Max = A
Else
    Max = B
End If
End Function
Function RevAy(Ay)
Dim O: O = Ay
Dim U&: U = UB(Ay)
Dim J&: For J = 0 To U
    O(J) = Ay(U - J)
Next
RevAy = O
End Function

Function Itr(Ay)
If Not IsEmpty(Ay) Then
    ChkIsAy Ay
End If
If Si(Ay) = 0 Then Set Itr = EmpColl Else Itr = Ay
End Function
Function EmpColl() As VBA.Collection
Static X As Collection, Y As Boolean
If Not Y Then Y = True: Set X = New Collection
Set EmpColl = X
End Function
Function IxzAy&(Ay, Ele)
Dim J&: For J = 0 To UB(Ay)
    If Ay(J) = Ele Then IxzAy = J: Exit Function
Next
IxzAy = -1
End Function
Function NotInAy(Ele, Ay) As Boolean
NotInAy = Not HasEle(Ay, Ele)
End Function
Function HasEle(Ay, M) As Boolean
Dim I: For Each I In Itr(Ay)
    If I = M Then HasEle = True: Exit Function
Next
End Function
Function IntersecSy(A$(), B$()) As String()
Dim I: For Each I In Itr(A)
    If HasEle(B, I) Then PushI IntersecSy, I
Next
End Function
Function SyzAp(ParamArray SyAp()) As String()
Dim I: For Each I In SyAp
    PushS SyzAp, CStr(I)
Next
End Function
Function AddSy(A$(), B$()) As String()
AddSy = AddAy(A, B)
End Function

Function AddSyAp(A$(), ParamArray SyAp()) As String()
Dim O$(): A = A
Dim IsY: For Each IsY In SyAp
    PushAy O, IsY
Next
End Function

Function AddAy(A, B)
AddAy = A
Dim I: For Each I In Itr(B)
    Push AddAy, I
Next
End Function
Function AddAyAp(Ay, ParamArray AyAp())
Dim O: O = Ay
Dim AyAv(): AyAv = AyAp
Dim IAy: For Each IAy In Itr(AyAv)
    PushAy O, IAy
Next
AddAyAp = O
End Function
Function EnsPth(Pth$)
If Dir(Pth, vbDirectory) = "" Then MkDir Pth
End Function

Sub D(Ay)
DmpAy Ay
End Sub
Sub DmpAy(Ay)
Dim V: For Each V In Itr(Ay)
    Debug.Print V
Next
End Sub
Sub CpyTblToFb(FmFb$, ToFb$, T)
Const CSub$ = "CpyTblToFb"
ChkFbNonLnkTblExist FmFb, T, CSub
ChkFbtNotExist ToFb, T, CSub
RunCQ FmtQQ("Select * into [?] in '?' from [?] in '?'", T, ToFb, T, FmFb)
End Sub
Sub ChkFbtNotExist(Fb$, T, Fun$)
If HasFbt(Fb, T) Then Thw FmtQQ(Fun & ": Table[?] already exist in Fb[?]", T, Fb)
End Sub

Sub ChkFbNonLnkTblExist(Fb$, T, Fun$)
Dim D As Database: Set D = Db(Fb)
If HasT(D, T) Then
    Dim CnStr$: CnStr = D.TableDefs(T).Connect
    If CnStr = "" Then Exit Sub
    Thw FmtQQ(Fun & ": Fb[?] should have non-Lnk-Tbl[?], but it has CnStr[?]", Fb, T, CnStr)
End If
Thw FmtQQ(Fun & ": Fb[?] should have Tbl[?]", Fb, T)
End Sub

Function HasRec(Q$) As Boolean
HasRec = Not CurrentDb.OpenRecordset(Q).EOF
End Function
Function CPth$()
CPth = Pth(CurrentDb.Name)
End Function
Function CPj() As VBProject: Set CPj = Vbe.ActiveVBProject: End Function
'---========================================================================================== AA_Ide_PjMode
Sub AA_Ide_PjMode(): End Sub
Function CPjMode() As vbext_VBAMode: CPjMode = CPj.Mode: End Function
Function IsBrkMode() As Boolean:       IsBrkMode = CPjMode = vbext_vm_Break: End Function
Function IsDesignMode() As Boolean: IsDesignMode = CPjMode = vbext_vm_Design: End Function
Function IsRunMode() As Boolean:       IsRunMode = CPjMode = vbext_vm_Design: End Function
'---========================================================================================== AA_Sts

Function CvFrm(A) As Access.Form
Set CvFrm = A
End Function
Sub SetAcsSts(Msg$)
If Msg = "" Then ClrAcsSts: Exit Sub
Application.SysCmd acSysCmdSetStatus, Msg
End Sub
Sub StsGen(Ffn$)
Sts "Generating file " & Ffn & "...."
End Sub

'---==========
Sub AA_Tmp(): End Sub
Function Tmp_QryNm$()
Tmp_QryNm = TmpNm("#Q")
End Function
Function NewQry(N$, Sql$) As DAO.QueryDef
Set NewQry = New QueryDef
NewQry.Name = N
NewQry.Sql = Sql
End Function


Sub DrpTmpA()
RunCQ "Drop Table [#A]"
End Sub
Function HasNm(Itr, Nm) As Boolean
Dim I: For Each I In Itr
    If I.Name = Nm Then HasNm = True: Exit Function
Next
End Function
Function HasTbl(T) As Boolean
CurrentDb.TableDefs.Refresh
HasTbl = HasNm(CurrentDb.TableDefs, T)
End Function

Function HasFbt(Fb$, T) As Boolean
HasFbt = HasT(Db(Fb), T)
End Function

Function HasT(Db As Database, T) As Boolean
HasT = HasNm(Db.TableDefs, T)
End Function

Function NoFbt(Fb$, T) As Boolean
NoFbt = Not HasFbt(Fb, T)
End Function

Sub Drp(T)
DbDrpT T
End Sub

Sub DrpTblAp(ParamArray TblAp())
Dim Av(): Av = TblAp
Dim T: For Each T In Av
    Drp T
Next
End Sub
Sub DrpTny(Tny$())
DbDrpTny Tny
End Sub

Sub DrpTT(TT$)
DbDrpTny Db, TT
End Sub
Sub DbDrpTT(D As Database, TT$)
For Each T In Termy(TT)
    DbDrpT D, T
Next
End Sub
Sub DltFfn(Ffn$)
If Dir(Ffn) = "" Then Exit Sub
On Error GoTo X
Kill Ffn
Exit Sub
X: Raise "Following file cannot be deleted.  Make sure it is not opened:" & vbCrLf & vbCrLf & Ffn, vbCritical
End Sub
Private Sub AlignLeft(Lo As ListObject, C$)
Lo.ListColumns(C).Total.HorizontalAlignment = XlHAlign.xlHAlignLeft
End Sub
Function LcRg(Lo As ListObject, C) As Range:    Set LcRg = Lc(Lo, C).Range: End Function
Function EntLc(Lo As ListObject, C) As Range:  Set EntLc = LcRg(Lo, C).EntireColumn: End Function
'---============================================================================================== AA_Xls_Lc
Sub AA_Xls_Lc(): End Sub
'---================================================================================================ AA_Xls_Lc_Fmt
Sub AA_Xls_Lc_Fmt(): End Sub



Function MinusSy(A$(), B$()) As String()
Dim IA: For Each IA In Itr(A)
    If Not HasEle(B, IA) Then Push MinusSy, IA
Next
End Function
Function AmQuoSq(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushS AmQuoSq, "[" & I & "]"
Next
End Function
Function MinusAy(A, B)
Dim O: O = A: Erase O
Dim IA: For Each IA In Itr(A)
    If Not HasEle(B, IA) Then Push O, IA
Next
MinusAy = O
End Function
Sub RunTrue2P(ShouldTrue As Boolean, Fun$, P1, P2)
If ShouldTrue Then Run Fun, P1, P2
End Sub
Sub RunTrue(ShouldTrue As Boolean, Fun$)
If ShouldTrue Then Run Fun
End Sub
Sub RunTrue1P(ShouldTrue As Boolean, Fun$, P)
If ShouldTrue Then Run Fun, P
End Sub
Function CvBtn(Ctl) As Access.CommandButton: Set CvBtn = Ctl: End Function
Function CvTBox(Ctl) As Access.CommandButton: Set CvTBox = Ctl: End Function
Sub SetCtlVis(Ctl, Vis As Boolean)
Select Case TypeName(Ctl)
Case "TextBox": CvTBox(Ctl).Visible = Vis
Case "CommandButton": CvBtn(Ctl).Visible = Vis
Case Else: Stop
End Select
End Sub
Function Opn_Rs(Sql$) As DAO.Recordset: Set Opn_Rs = CurrentDb.OpenRecordset(Sql): End Function
Function RplQ$(QStr$, By): RplQ = Replace(QStr, "?", By): End Function
Sub RfhWb(Wb As Workbook)
'Aim: Use current mdb as source to refresh given {pWorkbooks} data.
MinvWb Wb
RfhWbLo Wb
RfhWbQt Wb
RfhWbPc Wb
RfhWbPt Wb
MiniWb Wb
End Sub
Private Sub RfhWbLo(Wb As Workbook)
Dim CnStr$: CnStr = CnnStr_Mdb(CurrentDb.Name)
Sts RfhMsg(Wb, "List objects")
Dim Ws As Worksheet: For Each Ws In Wb.Worksheets
    Dim Lo As ListObject: For Each Lo In Ws.ListObjects
        Dim Qt As Excel.QueryTable
        Set Qt = Lo.QueryTable
        Qt.Connection = CnStr
        Qt.Refresh BackgroundQuery:=False
        DoEvents
    Next
Next
End Sub
Private Sub RfhWbQt(Wb As Workbook)
Dim CnStr$: CnStr = CnnStr_Mdb(CurrentDb.Name)
Sts RfhMsg(Wb, "Query tables")
Dim Ws As Worksheet: For Each Ws In Wb.Worksheets
    Dim Qt As Excel.QueryTable: For Each Qt In Ws.QueryTables
        Qt.Connection = CnStr
        Qt.Refresh False
    Next
Next
End Sub
Private Sub RfhWbPc(Wb As Workbook)
Dim CnStr$: CnStr = CnnStr_Mdb(CurrentDb.Name)
Sts RfhMsg(Wb, "Pivot caches")
Dim Pc As Excel.PivotCache: For Each Pc In Wb.PivotCaches
    If Pc.SourceType <> xlDatabase Then
        Pc.Connection = CnStr
        Pc.BackgroundQuery = False
        Pc.MissingItemsLimit = xlMissingItemsNone
    End If
    Pc.MissingItemsLimit = xlMissingItemsNone
    Pc.Refresh
Next
End Sub
Sub RfhWbPt(Wb As Workbook)
Sts RfhMsg(Wb, "Pivot tables")
Dim Ws As Worksheet: For Each Ws In Wb.Worksheets
    Dim Pt As Excel.PivotTable: For Each Pt In Ws.PivotTables
        Pt.RefreshTable
    Next
Next
End Sub
Private Function RfhMsg$(Wb As Workbook, Itm$)
RfhMsg = "RfhWb Wb[" & Wb.Name & "] is refreshing " & Itm & "......"
End Function
'---==================================================================================== A_Xls_Fmt_Rg
Sub BdrAround(pRge As Range)
 'mRge.BorderAround XlLineStyle.xlContinuous, xlMedium, ThemeColor:=2
With pRge.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 1
    .TintAndShade = 0
    .Weight = xlMedium
End With
With pRge.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 1
    .TintAndShade = 0
    .Weight = xlMedium
End With
With pRge.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 1
    .TintAndShade = 0
    .Weight = xlMedium
End With
With pRge.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 1
    .TintAndShade = 0
    .Weight = xlMedium
End With
End Sub
'---==================================================================================== A_Xls
Private Sub AA_Xls(): End Sub
Sub QuitWb(Wb As Workbook)
Wb.Application.Quit
End Sub
Function NwXls() As Excel.Application
Set NwXls = New Excel.Application
NwXls.WindowState = xlMinimized
NwXls.Visible = True
NwXls.DisplayAlerts = False
End Function
'---============================================================================================== A_IsXxxFn
Sub AA_AyFun(): End Sub
Sub PushObj(OAy, M)
Dim N&: N = Si(OAy)
ReDim Preserve OAy(N)
Set OAy(N) = M
End Sub
Sub PushS(O$(), S)
Dim N%: N = Si(O)
ReDim Preserve O(N)
O(N) = S
End Sub
Function Si&(Ay)
On Error Resume Next
Si = UBound(Ay) + 1
End Function
Function UB&(Ay)
UB = Si(Ay) - 1
End Function

'---=========================================================================== B_PfxSfx
'Nz***============================================================================== B_Xls
Sub AA_Xls_Fun(): End Sub
Function WsnzLo$(Lo As ListObject)
WsnzLo = WszLo(Lo).Name
End Function
Function WszLo(Lo As ListObject) As Worksheet
Set WszLo = Lo.Parent
End Function
Function WbzFx(Fx$) As Workbook
Set WbzFx = NwXls.Workbooks.Open(Fx)
End Function


'---=============================
Sub DmpOHMaxOrd()
DmpMaxOrdinalPosition ("@OH")
End Sub
Sub DmpMaxOrdinalPosition(T)
Debug.Print "FldCnt=" & FldCnt(T)
Debug.Print "MaxOrdinalPosition=" & MaxOrdinalPosition(T)
End Sub
Function FldCnt%(T)
FldCnt = CDb.TableDefs(T).Fields.Count
End Function

Private Function AA_Rseq(T, OldNew$()) As String()
AA_Rseq = Rseq(AA_Old(OldNew), Fny(T))
End Function
Private Function AA_ONew(OldNew$()) As TRst()
AA_ONew = TRst_Whe_EmpRst(TRstAy(OldNew))
End Function
Private Function AA_Old(OldNew$()) As String()
AA_Old = T1Ay(TRstAy(OldNew))
End Function

Function Rseq(Nseq$(), Oseq$()) As String()
ChkSubAy Nseq, Oseq
Rseq = AddSy(Nseq, MinusSy(Oseq, Nseq))
End Function

'---===============================
Function IsGoodDb(D As Database) As Boolean
On Error GoTo X
Dim A$: A = D.Name
IsGoodDb = True
Exit Function
X:
End Function

Function FldSno%(T, F)
FldSno = CurrentDb.TableDefs(T).Fields(F).OrdinalPosition
End Function

Function HasFld(T, F) As Boolean
Dim Fd As DAO.Field: For Each Fd In Td(T).Fields
    If Fd.Name = F Then HasFld = True: Exit Function
Next
End Function

Function Tny() As String()
Tny = TnyzDb(CurrentDb)
End Function

Function TnyzDb(D As Database) As String()
TnyzDb = Itn(D.TableDefs)
End Function

Function Fny(T) As String()
Dim A As DAO.TableDef: Set A = Td(T)
Fny = FnyzF(A.Fields)
End Function
Function FnyzF(F As DAO.Fields) As String()
FnyzF = Itn(F)
End Function
Function Itn(Itr) As String()
Dim I: For Each I In Itr
    PushS Itn, I.Name
Next
End Function
Sub ChkSubAy(SubAy, Super)
Const CSub$ = CMod & "ChkSubAy"
If Not IsSubAy(SubAy, Super) Then
    ThwPmvEr CSub, "Excess-In-SubAy", Join(MinusAy(SubAy, Super), vbCrLf), "Not SubAy"
End If
End Sub
Function IsSubAy(SubAy, Super) As Boolean
Dim ISub: For Each ISub In Itr(SubAy)
    If Not HasEle(Super, ISub) Then Exit Function
Next
IsSubAy = True
End Function

'---=========================================================================== B_Aft
Function AftSpc$(S)
AftSpc = Aft(S, " ")
End Function
Function Aft$(S, C$)
Dim P&: P = InStr(S, C)
If P = 0 Then ThwPmvEr "C", C, "not found in S[" & S & "]"
Aft = Mid(S, P + 1)
End Function
Function BefOrAll$(S, C$)
Dim P&: P = InStr(S, C)
If P = 0 Then BefOrAll = S: Exit Function
BefOrAll = Left(S, P - 1)
End Function

'---=========================================================================================== SqFun
Function SngEleSq(Ele) As Variant()
Dim O(): ReDim O(1 To 1, 1 To 1)
O(1, 1) = Ele
SngEleSq = O
End Function
'---=========================================================================================== RgFun
Function SqzRg(Rg As Range) As Variant()
If Rg.Count = 1 Then SqzRg = SngEleSq(Rg.Value): Exit Function
SqzRg = Rg.Value
End Function
Function CvRg(A) As Range
Set CvRg = A
End Function
Function C12zRg(Rg As Range) As C12
C12zRg = C12(Rg.Column, C2zRg(Rg))
End Function
Function NxtColCell(Rg As Range) As Range
Set NxtColCell = RgRC(Rg, 1, Rg.Column + 1)
End Function
Function LasColCell(Rg As Range) As Range
Set LasColCell = RgRC(Rg, 1, Rg.Columns.Count)
End Function
Function RgRC(Rg As Range, R, C) As Range
Set RgRC = Rg.Cells(R, C)
End Function

Function MaxCno(Ws As Worksheet)
MaxCno = Ws.Columns.Count
End Function

Function WszRg(Rg As Range) As Worksheet
Set WszRg = Rg.Parent
End Function
Function RgR(Rg As Range, R) As Range
Dim A As RCC
With A
    .R = Rg.Row + R - 1
    .C1 = Rg.Column
    .C2 = C2zRg(Rg)
    Set RgR = WsRCC(WszRg(Rg), .R, .C1, .C2)
End With
End Function
Function IsXlsxFn(Fn) As Boolean:  IsXlsxFn = HasSfx(Fn, ".xlsx"): End Function

'---=====================================================//
Function CntT&(T)
CntT = VzQ("Select Count(*) from [" & T & "]")
End Function
Function VzCQ(Sql)
VzCQ = CurrentDb.OpenRecordset(Sql).Fields(0).Value
End Function
Sub PromptQ(Sql, Optional Pfx$)
MsgBox Pfx & VzQ(Sql)
End Sub
Function SyzSql(Sql$) As String()
SyzSql = SyzRs(RszSql(Sql))
End Function
'------------------------------------------

Function SyzSS(SS$) As String()
SyzSS = Split(RmvDblSpc(Trim(SS)), " ")
End Function

Function FstWs(Wb As Workbook) As Worksheet
Set FstWs = Wb.Sheets(1)
End Function


Function ULin$(S, Optional ULinChr$ = "=")
ULin = String(Len(S), ULinChr)
End Function


#End If
