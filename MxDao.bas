Attribute VB_Name = "MxDao"
Option Compare Text
Option Explicit
Const CNs$ = "Dao"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxDao."
Function DbEng() As DBEngine
Set DbEng = DAO.DBEngine
End Function

Function CvCn(A) As ADODB.Connection
Set CvCn = A
End Function

Sub RplLoCnByFbt(Lo As ListObject, Fb, T)
With Lo.QueryTable
    Rpl_Wc_ByFb .Connection, Fb '<==
    .CommandType = xlCmdTable
    .CommandText = T '<==
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnINF = True
    .ListObject.DisplayName = Lon(T) '<==
    .Refresh BackgroundQuery:=False
End With
End Sub
Function TmpInpTny(D As Database) As String()
TmpInpTny = AwPfx(Tny(D), "#I")
End Function


Function NwWbzDbtmpInp(D As Database) As Workbook
Set NwWbzDbtmpInp = NwWbzDbtny(D, TmpInpTny(D))
End Function

Function NwWbzDbtny(D As Database, Tny$()) As Workbook
Dim T, O As Workbook
Set O = NwWb
For Each T In Itr(Tny)
    AddWszDbt O, D, CStr(T)
Next
DltSheet1 O
Set NwWbzDbtny = O
End Function
Function WbzFb(Fb) As Workbook
Dim D As Database: Set D = Db(Fb)
Set WbzFb = VisWb(NwWbzDbtny(D, Tny(D)))
End Function

Function SetWsn(Ws As Worksheet, Nm$) As Worksheet
Const CSub$ = CMod & "SetWsn"
Set SetWsn = Ws
If Nm = "" Then Exit Function
If Ws.Name = Nm Then Exit Function
If HasWs(WbzWs(Ws), Nm) Then
    Dim Wb As Workbook: Set Wb = WbzWs(Ws)
    Thw CSub, "Wsn exists in Wb", "Wsn Wbn Wny-in-Wb", Nm, Wbn(Wb), Wny(Wb)
End If
Ws.Name = Nm
End Function
Sub NwLozFbtzSamp()
'    Application.CutCopyMode = False
'    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
'        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=C:\Users\user\Desktop\SAPAccessReports\DutyPrepay5\DutyP" _
'        , _
'        "repay5_Data.mdb;Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:D" _
'        , _
'        "atabase Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Glob" _
'        , _
'        "al Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=Fals" _
'        , _
'        "e;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Supp" _
'        , _
'        "ort Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceFiel" _
'        , "d Validation=False"), Destination:=Range("$H$4")).QueryTable
'        .CommandType = xlCmdTable
'        .CommandText = Array("@RptM")
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = True
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = False
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .PreserveColumnInfo = True
'        .SourceDataFile = _
'        "C:\Users\user\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
'        .ListObject.DisplayName = "Table_DutyPrepay5_Data_1"
'        .Refresh BackgroundQuery:=False
'    End With

End Sub
Function NwLozFbt(At As Range, Fb, T) As ListObject
Dim Ws As Worksheet: Set Ws = WszRg(At)
Dim Lo As ListObject: Set Lo = Ws.ListObjects.Add(xlSrcExternal, OleCnStrzFb(Fb), Destination:=At)
Dim Qt As QueryTable: Set Qt = Lo.QueryTable
With Qt
    .CommandType = xlCmdTable
    .CommandText = T
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = False
    .RefreshPeriod = 0
    .PreserveColumnInfo = True
    .ListObject.DisplayName = LonzT(T)
    .Refresh BackgroundQuery:=False
End With
AutoFitLo Lo
End Function

Sub AutoFitLo(A As ListObject)
A.DataBodyRange.EntireColumn.AutoFit
End Sub

Function AddWszDbt(Wb As Workbook, Db As Database, T, Optional Wsn0$, Optional AddgWay As eWsdAddWay) As Worksheet
Dim O As Worksheet: Set O = AddWs(Wb, StrDft(Wsn0, T))
Dim A1 As Range: Set A1 = A1zWs(O)
NwLozDbt Db, T, A1, AddgWay
End Function

Function NwWbzFbOup(Fb, Optional AddgWay As eWsdAddWay) As Workbook '
Dim O As Workbook, D As Database
Set O = NwWb
Set D = Db(Fb)
AddWszDbTny O, D, OupTny(D), AddgWay
DltWsIf O, "Sheet1"
Set NwWbzFbOup = O
End Function

Function NwWbzDbt(D As Database, T, Optional Wsn$ = "Data") As Workbook
Set NwWbzDbt = WszRg(AddWszDbt(NwWb, D, T, Wsn))
End Function

Sub ClrLo(A As ListObject)
If A.ListRows.Count = 0 Then Exit Sub
A.DataBodyRange.Delete xlShiftUp
End Sub

Sub SetQtFbt(Qt As QueryTable, Fb, T)
With Qt
    .CommandType = xlCmdTable
    .Connection = OleCnStrzFb(Fb) '<--- Fb
    .CommandText = T '<-----  T
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnINF = True
    .Refresh BackgroundQuery:=False
End With
End Sub
Sub PutFbtAt(Fb, T$, At As Range, Optional Lon0$)
Dim O As ListObject
Set O = WszRg(At).ListObjects.Add(SourceType:=XlSourceType.xlSourceWorkbook, Destination:=At)
SetLon O, Dft(Lon0, Lon(T))
SetQtFbt O.QueryTable, Fb, T
End Sub
Sub FxzTny(Fx, Db As Database, Tny$())
NwWbzDbtny(Db, Tny).SaveAs Fx
End Sub

Function NwWs_FmDbt(D As Database, T, Optional Wsn$) As Worksheet
Dim Sq(): Sq = SqzT(D, T)
Dim A1 As Range: Set A1 = NwA1(Wsn)
Set NwWs_FmDbt = WszLo(NwLozSq(Sq(), A1))
End Function
