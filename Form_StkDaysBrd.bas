VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_StkDaysBrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = CLib & "Form_StkDaysBrd."
Option Base 0
Private Sub CmdExit_Click()
DoCmd.Close
End Sub
Private Sub Form_BeforeUpdate(Cancel As Integer)
Me.DteUpd.Value = Now()
End Sub
Private Sub Form_Open(Cancel As Integer)
DoCmd.Maximize
'Load list of string of YYYY/MM/DD to Me.xDate, load at last 13 months of date (Inlude current month).  If a month has a date of OH, use the max date of that month.  If month has no OH, use the last date of that month.
DrpCT "#Tmp"
RunCQ "Create table `#Tmp` (Dte Date)"
Dim mYY As Byte, mMM As Byte, mDD As Byte
mYY = Year(Date) - 2000
mMM = Month(Date)
Dim mDte As Date
Dim J%
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.TableDefs("#Tmp").OpenRecordset
For J% = 0 To 12
    With CurrentDb.OpenRecordset(FmtStr("SELECT Max(x.DD) as DD FROM OH x where YY={0} and MM={1}", mYY, mMM))
        If IsNull(!DD) Then
            mDte = LasDtezYM(mYY, mMM)
        Else
            mDte = CDate(mYY + 2000 & "/" & mMM & "/" & !DD)
        End If
        .Close
    End With
    With mRs
        .AddNew
        !Dte = mDte
        .Update
    End With
    If mMM = 1 Then
        mMM = 12
        mYY = mYY - 1
    Else
        mMM = mMM - 1
    End If
Next
Dim mA$: mA = ""
With CurrentDb.OpenRecordset("Select Distinct Dte from `#Tmp` order by Dte Desc")
    While Not .EOF
        If mA = "" Then
            mA = Format(!Dte, "YYYY/MM/DD")
        Else
            mA = mA & ";" & Format(!Dte, "YYYY/MM/DD")
        End If
        .MoveNext
    Wend
    .Close
End With
Me.xDate.RowSourceType = "Value List"
Me.xDate.RowSource = mA
Me.xDate.SetFocus
End Sub
Private Sub xDate_AfterUpdate()
'StkDays: CdSHBrand: YY MM DD CdSHBrand StkDays Rmk DteCrt DteUpd

Dim mYY As Byte: mYY = Val(Left(Me.xDate.Value, 4)) - 2000
Dim mMM As Byte: mMM = Val(Mid(Me.xDate.Value, 6, 2))
Dim mDD As Byte: mDD = Val(Right(Me.xDate.Value, 2))

DrpCT "#Tmp"
RunCQ "Select Distinct CdSHBrand into [#Tmp] from (OH x INNER JOIN q1SKU a on x.SKU=a.SKU) INNER JOIN SHBrandQ b on a.CdB=b.CdQly"
RunCQ "Alter Table `#Tmp` add column YY Byte, MM Byte, DD Byte"
RunCQ FmtStr("Update `#Tmp` set YY={0},MM={1},DD={2}", mYY, mMM, mDD)
RunCQ FmtStr("Insert into StkDays (YY,MM,DD,CdSHBrand) Select {0},{1},{2},x.CdSHBrand From  `#Tmp` x LEFT JOIN StkDays a on x.YY=a.YY and x.MM=a.MM and x.DD=a.DD and x.CdSHBrand=a.CdSHBrand where a.CdSHBrand is null", mYY, mMM, mDD)

Me.RecordSource = "SELECT CdCo,x.*, NmSHBrand, NmFinStream" & _
" FROM ((StkDays x LEFT JOIN SHBrand a ON x.CdSHBrand = a.CdSHBrand) LEFT JOIN FinStream b ON a.CdFinStream = b.CdFinStream) Left jOin Co c on c.Co=x.Co" & _
OH3YmdBexp(mYY, mMM, mDD) & _
" Order by CdCo, NmFinStream, NmSHBrand"
End Sub
