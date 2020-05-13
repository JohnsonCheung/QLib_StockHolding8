Attribute VB_Name = "MxCliMHDFb"
Option Explicit
Option Compare Text
Const CNs$ = "Mhd.Fb"
Const CLib$ = "QApp."
Const CMod$ = CLib & "MxCliMHDFb."

Const L$ = "C:\Users\Public\Logistic\"
Public Const ArHom$ = "C:\Users\Public\DebtorAging4 and ARStmt\"
Public Const LgsHom$ = "C:\Users\Public\Logistic\"
Public Const LgsDtaHom$ = LgsHom & "SapData\"
Public Const MhdAppnn$ = "Aging EStmt CrRel Duty StkHld TaxAlert TaxCmp RelCst CrRvw"
Public Const DutyPth$ = LgsHom & "DutyPrepay7\"
Public Const TaxCmpPth$ = LgsHom & "TaxCmp\"
Public Const TaxAlertPth$ = LgsHom & "TaxAlert\"
Public Const RelCstPth$ = LgsHom & "RelCst\"
Public Const StkHld8Pth$ = LgsHom & "StockHolding8\"
Public Const StkHld8TpPth$ = StkHld8Pth & "WorkingDir\Templates\"
'------------------------------------------
Public Const AgingFba$ _
                            = ArHom & "ARStmt(eStmt)\ARStmt.accdb"
Public Const AgingDtaFb$ _
                            = ArHom & "ARStmt(eStmt)\ARStmt_Data.accdb"
Public Const EStmtFba$ _
                            = ArHom & "ARStmt(eStmt)\ARStmt.accdb"
Public Const EStmtDtaFb$ _
                            = ArHom & "ARStmt(eStmt)\ARStmt_Data.accdb"
Public Const DutyDtaFb$ _
                            = DutyPth & "DutyPrepay7_Data.accdb"
Public Const DutyFba$ _
                            = DutyPth & "DutyPrepay7.accdb"
Public Const TaxCmpFba$ _
                            = TaxCmpPth & "TaxCmp v1.3.accdb"
Public Const TaxAlertFba$ _
                            = TaxAlertPth & "TaxAlert 1.4\TaxAlert 1.4.accdb"
Public Const RelCstFba$ _
                            = RelCstPth & "RelCst 1.0\RelCst 1.0.accdb"
Public Const StkHld8Fba$ _
                            = StkHld8Pth & "StockHolding8.accdb"
Public Const StkHld8MB52Tp$ _
                            = StkHld8TpPth & "Stock Holding Template.xlsx"
Public Const StkHld8DtaFb$ _
                            = StkHld8Pth & "StockHolding8_Data.accdb"
Public Const CrRlsFba$ _
                            = ArHom & "CrHldRls2\CrHldRls2.accdb"
Public Const SalTxtFx$ _
                            = LgsDtaHom & "Sales Text.xlsx"

Function StkHld8TmpFba$()
StkHld8TmpFba$ = TmpHom & "TmpStockHolding8.accdb"
End Function

':DtaFb: :Fb #Data-Fb#
Function MhdDtaFbAy() As String()
Dim O$()
PushI O, StkHld8DtaFb
PushI O, DutyDtaFb
PushI O, EStmtDtaFb
PushI O, AgingDtaFb
MhdDtaFbAy = O
End Function

Function MhdFbaAy() As String()
Dim O$()
PushI O, StkHld8Fba
PushI O, DutyFba
PushI O, TaxAlertFba
PushI O, TaxCmpFba
PushI O, RelCstFba
PushI O, EStmtFba
PushI O, AgingFba
PushI O, CrRlsFba
MhdFbaAy = O
End Function

Property Get MhdAppFbDic() As Dictionary
Const A$ = "N:\SAPAccessReports\"
BfrClr
BfrV "Duty     " & A & "DutyPrepay\.accdb"
BfrV "SkHld    " & A & "StkHld\.accdb"
BfrV "ShpRate  " & A & "DutyPrepay\StockShipRate_Data.accdb"
BfrV "ShpCst   " & A & "StockShipCost\.accdb"
BfrV "TaxCmp   " & A & "TaxExpCmp\.accdb"
BfrV "TaxAlert " & A & "TaxRateAlert\.accdb"
Set MhdAppFbDic = Dic(BfrLy)
End Property
