Attribute VB_Name = "JMxRsPrp"
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "JMxRsPrp."
#If False Then
Option Explicit
Function FnyzRs(Rs As Recordset) As String()
FnyzRs = Itn(Rs.Fields)
End Function
#End If
