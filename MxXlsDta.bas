Attribute VB_Name = "MxXlsDta"
Option Explicit
Option Compare Text
Const CNs$ = "Xls.Dta"
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsDta."

Function DrszFxq(Fx, Q) As Drs
DrszFxq = DrszArs(CnzFx(Fx).Execute(Q))
End Function

Private Sub DrszFxq__Tst()
Dmp WnyzFx(SalTxtFx)
BrwDrs DrszFxw(SalTxtFx, "Sheet1")
End Sub

Function WszDt(A As Dt) As Worksheet
Dim O As Worksheet
Set O = NwWs(A.DtNm)
LozDrs DrszDt(A), A1zWs(O)
WszDt = O
End Function
