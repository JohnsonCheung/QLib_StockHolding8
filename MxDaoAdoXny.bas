Attribute VB_Name = "MxDaoAdoXny"
Option Explicit
Option Compare Text
Function TnyzCat(A As Catalog) As String()
TnyzCat = Itn(A.Tables)
End Function

Function TnyzFb(Fb) As String()
TnyzFb = Tny(Db(Fb))
End Function

Function TnyzFbByAdo(Fb) As String()
TnyzFbByAdo = AeKss(TnyzCat(CatzFb(Fb)), "MSys* f_*_Data")
End Function

Function Wny(Fx, Optional InlAllOtherTbl As Boolean) As String()
Wny = WnyzFx(Fx)
End Function
Function TnyzFx(Fx) As String()
TnyzFx = Itn(AxTdszFx(Fx))
End Function

'--
Private Sub WnyzFx__Tst()
Dim Fx$
GoSub Z
'GoSub T1
'GoSub T2
Exit Sub
Tst:
    Act = WnyzFx(Fx)
    C
    Return
T1:
    Fx = SalTxtFx
    Ept = SyzSS("")
    GoTo Tst
T2:
    Fx = "C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
    Ept = SyzSS("")
    GoTo Tst
Z:
    DmpAy WnyzFx(SalTxtFx)
    Return
End Sub
Function WnyzFx(Fx, Optional InlAllOtherTbl As Boolean) As String()
Const CSub$ = CMod & "WnyzFx"
ChkFfnExist Fx, CSub, "Fx"
Dim Tny$(), T
Tny = TnyzCat(CatzFx(Fx))
If InlAllOtherTbl Then
    WnyzFx = Tny
    Exit Function
End If
For Each T In Itr(Tny)
    PushNB WnyzFx, WsnzCattn(T)
Next
End Function
'--
Function WsnzCattn$(Cattn)
':Cattn: :TblNm ! #Cat-Tbl-Nm#
If HasSfx(Cattn, "FilterDatabase") Then Exit Function
WsnzCattn = RmvSfx(RmvSngQuo(Cattn), "$")
End Function

Function FFzFxw$(Fx, Optional W$)
FFzFxw = Tml(FnyzFxw(Fx, W))
End Function

Function FnyzAfds(A As ADODB.Fields) As String()
FnyzAfds = Itn(A)
End Function

'**Fny
Function FnyzArs(A As ADODB.Recordset) As String(): FnyzArs = FnyzAfds(A.Fields): End Function
Function FnyzAxTd(T As Adox.Table) As String():    FnyzAxTd = Itn(T.Columns):     End Function
Function FnyzFbt(Fb, T) As String():                FnyzFbt = Fny(Db(Fb), T):     End Function
Function FnyzFbtAdo(Fb, T) As String()
Dim C As Adox.Catalog
Set C = CatzFb(Fb)
FnyzFbtAdo = FnyzAxTd(C.Tables(T))
End Function
