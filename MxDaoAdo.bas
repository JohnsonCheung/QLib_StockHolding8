Attribute VB_Name = "MxDaoAdo"
Option Compare Text
Option Explicit
#If Doc Then
'Ax:Cml #AdoX#
'Td:Cml #Table-Definition#
#End If
Const CLib$ = "QDao."
Const CNs$ = "Ado"
Const CMod$ = CLib & "MxDaoAdo."
Function AxTd(C As Catalog, T) As Adox.Table  'Ado.Table definition
Set AxTd = C.Tables(T)
End Function

Function AxTbn$(Wsn)  ' #Adox-Table-Name# format for Wsn: When Wsn IsNm, just add Sfx-$, otherwise SngQuo(Add Sfx-$)
AxTbn = IIf( _
    IsNm(Wsn), _
        Wsn & "$", _
        QuoSng(Wsn & "$"))
End Function

Function AxTdszFx(Fx) As Adox.Tables
Dim C As Catalog: Set C = CatzFx(Fx) ' it is must.  CatzFx(Fx).Tables does not work. Because in CatzFx(Fx) will be disposed before it can return .Table
Set AxTdszFx = C.Tables
End Function

Function Ars(Cn As ADODB.Connection, Q) As ADODB.Recordset: Set Ars = ArszCnq(Cn, Q): End Function
Function ArszCnq(Cn As ADODB.Connection, Q) As ADODB.Recordset: Set ArszCnq = Cn.Execute(Q): End Function

'**ArsCol
Function ColzArs(A As ADODB.Recordset, Optional F = 0) As Variant():     ColzArs = ZZIntozArs(EmpAv, A, F): End Function
Function IntAyzArs(A As ADODB.Recordset, Optional F = 0) As Integer(): IntAyzArs = ZZIntozArs(EmpIntAy, A, F): End Function

'**Ars
Function ArszFxQ(Fx$, Q$) As ADODB.Recordset:                Set ArszFxQ = CnzFx(Fx).Execute(Q):                 End Function
Function ArszFxw(Fx$, W, Optional Bexp$) As ADODB.Recordset: Set ArszFxw = ArszFxQ(Fx, SqlSelStar_Fm(AxTbn(W))): End Function

'**Arun
Sub ArunzFbQ(Fb, Q): CnzFb(Fb).Execute Q: End Sub

'**AxFun-Cat
Function Cat(A As ADODB.Connection) As Catalog
Set Cat = New Catalog
Set Cat.ActiveConnection = A
End Function
Function CatzFb(Fb) As Catalog: Set CatzFb = Cat(CnzFb(Fb)): End Function
Function CatzFx(Fx) As Catalog: Set CatzFx = Cat(CnzFx(Fx)): End Function

'**AxFun-Cn
Function CnzFb(A) As ADODB.Connection: Set CnzFb = Cn(AdoCnStrzFb(A)): End Function
Function CnzFx(Fx) As ADODB.Connection: ChkFfnExist Fx, "CnzFx", "Excel file": Set CnzFx = Cn(AdoCnStrzFx(Fx)): End Function
Function Cn(AdoCnStr) As ADODB.Connection: Set Cn = New ADODB.Connection: Cn.Open AdoCnStr: End Function

Function AdoCnStrzFb$(A)
'Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserINF Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
'AdoCnStrzFb = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;", A)
Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;User ID=Admin;Mode=Share Deny None;"
'Locking Mode=1 means page (or record level) according to https://www.spreadsheet1.com/how-to-refresh-pivottables-without-locking-the-source-workbook.html
'The ADO connection object initialization property which controls how the database is locked, while records are being read or modified is: Jet OLEDB:Database Locking Mode
'Please note:
'The first user to open the database determines the locking mode to be used while the database remains open.
'A database can only be opened is a single mode at a time.
'For Page-level locking, set property to 0
'For Row-level locking, set property to 1
'With 'Jet OLEDB:Database Locking Mode = 0', the source spreadshseet is locked, while PivotTables update. If the property is set to 1, the source file is not locked. Only individual records (Table rows) are locked sequentially, while data is being read.
AdoCnStrzFb = FmtQQ(C, A)
End Function

Function AdoCnStrzFx$(A)
'AdoCnStrzFx = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?", A) 'Try
AdoCnStrzFx = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""", A) 'Ok
End Function

'**AxCv
Function CvAdoTy(A) As ADODB.DataTypeEnum:    CvAdoTy = A: End Function
Function CvAxTd(A) As Adox.Table:          Set CvAxTd = A: End Function ' #Catalog-Table#

Function DftTny(Tny0, Fb) As String()
If IsMissing(Tny0) Then
    DftTny = TnyzFb(Fb)
Else
    DftTny = CvSy(Tny0)
End If
End Function

Function DrzAfds(A As ADODB.Fields, Optional N%) As Variant()
Dim F As ADODB.Field
For Each F In A
   PushI DrzAfds, F.Value
Next
End Function

Function VzScvl(Scvl$, Nm$)
VzScvl = IsBet(EnsSfx(Scvl, ";"), Nm & "=", ";")
End Function

Function DtaSrczScvl(Scvl$)
DtaSrczScvl = VzScvl(Scvl, "Data Source")
End Function

Function AxTyStr$(T As Adox.DataTypeEnum)
Select Case True
Case True
End Select
Stop
End Function

Function DyzArs(A As ADODB.Recordset) As Variant()
While Not A.EOF
    PushI DyzArs, DrzAfds(A.Fields)
    A.MoveNext
Wend
End Function

Function FnyzFxw(Fx, Optional W$ = "") As String() ' ret Fny of Fx->W.  If W is blnk, fst W.
Dim C As Catalog: Set C = CatzFx(Fx)
Dim T As Adox.Table: Set T = C.Tables(AxTbn(W))
FnyzFxw = FnyzAxTd(T)
End Function

Function HasFbt(Fb, T) As Boolean:   HasFbt = HasEle(TnyzFb(Fb), T): End Function
Function HasFxw(Fx, Wsn) As Boolean: HasFxw = HasEle(WnyzFx(Fx), Wsn): End Function


Sub RunCnSqy(Cn As ADODB.Connection, Sqy$())
Dim Q
For Each Q In Itr(Sqy)
   Cn.Execute Q
Next
End Sub

Function SyzArs(A As ADODB.Recordset, Optional Col = 0) As String()
SyzArs = IntoColzArs(EmpSy, A, Col)
End Function

Private Sub ArunzFbQ__Tst()
Const Fb$ = DutyFba
Const Q$ = "Select * into [#a] from Permit"
DrpFbt Fb, "#a"
ArunzFbQ Fb, Q
End Sub

Private Sub Cn__Tst()
Dim O As ADODB.Connection
Set O = Cn(GetCnStr_ADO_SampSQL_EXPR_NOT_WRK)
Stop
End Sub

Private Sub AdoCnStrzFb__Tst()
Dim CnStr$
'
CnStr = AdoCnStrzFb(DutyDtaFb)
GoSub Tst
'
CnStr = AdoCnStrzFb(CurrentDb.Name)
'GoSub Tst
Exit Sub
Tst:
    Cn(CnStr).Close
    Return
End Sub

Private Sub CnzFb__Tst()
Dim Cn
Set Cn = CnzFb(DutyDtaFb)
Stop
End Sub

Private Sub DrszCnq__Tst()
Dim Cn As ADODB.Connection: Set Cn = CnzFx(SalTxtFx)
Dim Q$: Q = "Select * from [Sheet1$]"
WszDrs DrszCnq(Cn, Q)
End Sub

Private Sub DrszFbqAdo__Tst()
Const Fb$ = DutyDtaFb
Const Q$ = "Select * from Permit"
BrwDrs DrszFbqAdo(Fb, Q)
End Sub

Private Sub DyzArs__Tst()
Dim S$
Const Q$ = "Select * from KE24"
S = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute S
BrwDy DyzArs(ArszCnq(CnzFb(DutyDtaFb), Q))
End Sub

'--
Function HasReczArs(A As ADODB.Recordset) As Boolean
HasReczArs = Not NoReczArs(A)
End Function

Function NoReczArs(A As ADODB.Recordset) As Boolean
NoReczArs = A.EOF And A.BOF
End Function

'== X
Private Function ZZIntozArs(Into, A As ADODB.Recordset, F)
Dim O: O = NwAy(Into)
With A
    While Not .EOF
        PushI O, Nz(A(F))
        .MoveNext
    Wend
End With
End Function
