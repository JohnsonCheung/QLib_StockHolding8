Attribute VB_Name = "MxDaoDefFds"
Option Compare Text
Option Explicit
Const CNs$ = "Def"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDefFds."

Function CslzFds$(A As DAO.Fields)
':Fds: :Dao.Fields
Stop
CslzFds = Csl(AvzItr(A))
End Function

Function CslzFdsFny$(A As DAO.Fields, Fny$())
':Fds: :Dao.Fields
CslzFdsFny = Csl(DrzFdsFny(A, Fny))
End Function

Function DrzFds(A As DAO.Fields, Optional FF$) As Variant()
DrzFds = DrzFdsFny(A, FnyzFFRs(FF, A))
End Function

Function DrzFdsFny(A As DAO.Fields, Fny$()) As Variant()
Dim F: For Each F In Fny
    PushI DrzFdsFny, Nz(A(F).Value)
Next
End Function

Private Sub DrzFds__Tst()
Dim Rs As DAO.Recordset, Dy()
Set Rs = RelCstPgmDb.OpenRecordset("Select * from YMGRnoIR")
With Rs
    While Not .EOF
        PushI Dy, DrzFds(Rs.Fields)
        .MoveNext
    Wend
    .Close
End With
BrwDy Dy
End Sub

Private Sub DrzFds1__Tst()
Dim Rs As DAO.Recordset, Dr(), D As Database
Set Rs = RszQ(DutyDtaDb, "Select * from SkuB")
With Rs
    While Not .EOF
        Dr = DrzRs(Rs)
        Debug.Print JnComma(Dr)
        .MoveNext
    Wend
    .Close
End With
End Sub
