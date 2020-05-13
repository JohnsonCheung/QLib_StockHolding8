Attribute VB_Name = "MxDaoAdoDta"
Option Explicit
Option Compare Text
Const CNs$ = "Ado.Dta"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoAdoDta."
Function DrszCnq(Cn As ADODB.Connection, Q) As Drs
DrszCnq = DrszArs(ArszCnq(Cn, Q))
End Function

Function DrszFbqAdo(Fb, Q) As Drs
DrszFbqAdo = DrszArs(ArszFbq(Fb, Q))
End Function

Function DrszArs(A As ADODB.Recordset) As Drs
DrszArs = Drs(FnyzArs(A), DyzArs(A))
End Function
