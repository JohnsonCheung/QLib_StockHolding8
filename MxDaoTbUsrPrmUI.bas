Attribute VB_Name = "MxDaoTbUsrPrmUI"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDaoTbUsrPrmUI."

Sub TglPm(D As Database, Pmn$): SetPmv D, Pmn, Not Pmv(D, Pmn): End Sub

Sub TglCPm(Pmn$): TglPm CDb, Pmn: End Sub

Sub SelCFxPm(Pmn$, Optional FxSpecDes$ = "Select a Xlsx file", Optional tboxReqry As Access.TextBox)
SelFxPm CDb, Pmn, FxSpecDes, tboxReqry
End Sub

Sub SelFxPm(D As Database, Pmn$, Optional FxSpecDes$ = "Select a Xlsx file", Optional tboxReqry As Access.TextBox)
Dim Fx$: Fx = SelFx(Nz(Pmv(D, Pmn), ""), FxSpecDes): If Fx <> "" Then SetPmv D, Pmn, Fx
If Not IsNothing(tboxReqry) Then
    tboxReqry.Requery
End If
End Sub

Sub SelCPthPm(Pmn$, Optional T As Access.TextBox)
SelPthPm CDb, Pmn, T
End Sub

Sub SelCPthPmTBox(T As Access.TextBox)
SelPthPm CDb, T.Name, T
End Sub

Sub SelPthPm(D As Database, Pmn$, Optional T As Access.TextBox)
Dim P$: P = SelPth(Nz(Pmv(D, Pmn), ""))
If P <> "" Then
    SetPmv D, Pmn, P
    If Not IsNothing(T) Then T.Requery
End If
End Sub

Sub SelPthCPm(Pmn$, Optional T As Access.TextBox)
SelPthPm CDb, Pmn, T
End Sub
