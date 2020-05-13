Attribute VB_Name = "MxDaoDbSetFv"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDbSetFv."

Sub SetVzQ(D As Database, Q, V)
VzRs(D.OpenRecordset(Q)) = V
End Sub

Sub SetVzRs(A As DAO.Recordset, V)
If NoRec(A) Then
    A.AddNew
Else
    A.Edit
End If
A.Fields(0).Value = V
A.Update
End Sub

Sub SetVzRsF(Rs As DAO.Recordset, Fld, V)
With Rs
    .Edit
    .Fields(Fld).Value = V
    .Update
End With
End Sub

Sub SetVzSsk(D As Database, T, F$, SskvSet(), V)
VzRs(Rs(D, SqlSel_F_T_WhF_Eq(F, T, SskFldn(D, T), SskvSet))) = V
End Sub
