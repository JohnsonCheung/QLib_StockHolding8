Attribute VB_Name = "MxDaoReadDr"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoReadDr."

Private Sub DrzQ__Tst()
MsgBox JnSpc(DrzQ(CDb, "Select YY,MM,DD from[@OH]where YY=20"))
End Sub

Function DrzQ(D As Database, Q) As Variant()
DrzQ = DrzRs(Rs(D, Q))
End Function
