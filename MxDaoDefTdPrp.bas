Attribute VB_Name = "MxDaoDefTdPrp"
Option Explicit
Option Compare Text
Const CNs$ = "Def"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoDefTdPrp."

Sub SetTblDeszDi(D As Database, TblDesDi As Dictionary)
Dim T: For Each T In TblDesDi.Keys
    SetTblDes D, T, TblDesDi(T)
Next
End Sub

Sub SetFldDeszDi(D As Database, FldDesDi As Dictionary)
Dim TF: For Each TF In FldDesDi.Keys
    Dim T$, F$
    With BrkDot(TF)
        T = .S1
        F = .S2
    End With
    SetFldDes D, T, F, FldDesDi(TF)
Next
End Sub

Function TbDesDi(D As Database) As Dictionary
Dim T, O As New Dictionary
For Each T In Tni(D)
    AddKvIfNB O, T, TblDes(D, T)
Next
Set TbDesDi = O
End Function

Sub SetTbPrp(D As Database, T, P$, V)

End Sub

Function DaoPvzP(P As DAO.Properties, Pn$)
If HasPrp(P, Pn) Then DaoPvzP = P(Pn).Value
End Function

Function DaoPrps(DaoPrpsObj) As DAO.Properties
Set DaoPrps = DaoPrpsObj.Properties
End Function

Function DaoPv(DaoPrpObj, P$)
DaoPv = DaoPvzP(DaoPrpObj.Properties, P)
End Function
