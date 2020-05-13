Attribute VB_Name = "gzMigrate"
Option Compare Text
Option Explicit
Const CMod$ = CLib & "gzMigrate."
#If False Then
Private Const Pgm20200215FbFn$ = "StockHolding8(2020-02-15).accdb"
Private Const PgmFn$ = "StockHolding8.accdb"
Private Const Hom$ = "C:\users\public\Logistic\StockHolding8\"
Private Const DtaFb$ = Hom & "StockHolding8_Data.accdb"
Private Const BkupPth$ = Hom & "Backup\"
Private Const Pgm20200215Fb$ = Hom & Pgm20200215FbFn
Private Const BkupPgm20200215Fb$ = BkupPth & Pgm20200215FbFn

Private Sub Pgm20200215Fb__Tst()
MsgBox HasFfn(Pgm20200215Fb)
MsgBox HasFfn(DtaFb)
End Sub

Sub MigrateTblUsrPrm()
Exit Sub
If HasFbt(DtaFb, "UsrPrm") Then Exit Sub
If NoFfn(Pgm20200215Fb) Then
    BrwAy ErMsg
    Quit
End If
CpyTblToFb Pgm20200215Fb, DtaFb, "UsrPrm"
EnsPth BkupPth
DltFfnIf BkupPgm20200215Fb
Fso.MoveFile Pgm20200215Fb, BkupPth
End Sub

Private Sub ErMsg__Tst()
BrwAy ErMsg
End Sub

Function ErMsg() As String()
Dim O$()
PushI O, "UnExpected Error -- Program cannot run"
PushI O, "======================================"
PushI O, "[Data accdb file] does not have table-[UsrPrm], and,"
PushI O, "[Old renamed pgm file] does not exist."
PushI O, ""
PushI O, FmtQQ("[Old renamed program file] = [?]", Pgm20200215Fb)
PushI O, FmtQQ("[Data accdb file]          = [?]", DtaFb)
PushI O, ""
PushI O, "[Old renamed pgm file] is renamed from the old program"
PushI O, FmtQQ("   in Path        [?]", Hom)
PushI O, FmtQQ("   From file name [?]", PgmFn)
PushI O, FmtQQ("   To   file name [?]", Pgm20200215FbFn)
PushI O, ""
PushI O, "Table-[UsrPrm] must exist in [Data accdb file] and it is copied from [Old renamed program file]."
PushI O, "No Table-[UsrPrm], the program cannot run."
ErMsg = O
End Function
#End If
