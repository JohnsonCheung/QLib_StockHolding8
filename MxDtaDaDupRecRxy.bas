Attribute VB_Name = "MxDtaDaDupRecRxy"
Option Explicit
Option Compare Text
Const CNs$ = "Dta.Drs"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaDupRecRxy."

Function DupRecRxy(Dy()) As Long()
Dim DupD(): DupD = DywDup(Dy)
Dim Dr, Ix&, O&()
For Each Dr In Dy
    If HasDr(DupD, Dr) Then PushI O, Ix
    Ix = Ix + 1
Next
If Si(O) < Si(DupD) * 2 Then Stop
DupRecRxy = O
End Function

Function DupRecRxyzColIx(Dy(), ColIx&) As Long()
Dim D As New Dictionary, FstIx&, V, O As New Dictionary, Ix&, I
For Ix = 0 To UB(Dy)
    V = Dy(Ix)(ColIx)
    If D.Exists(V) Then
        O.PushParChd V, D(V)
        O.PushParChd V, Ix
    Else
        D.Add V, Ix
    End If
Next
For Each I In O.ParAetzRel.Itms
    PushIAy DupRecRxyzColIx, O.ParChd(I).Av
Next
End Function

Function DupRecRxyzFF(A As Drs, FF$) As Long()
Dim Fny$(): Fny = Termy(FF)
If Si(Fny) = 1 Then
    DupRecRxyzFF = IxyzDup(ColzDrs(A, Fny(0)))
    Exit Function
End If
Dim ColIxy&(): ColIxy = Ixy(A.Fny, Fny)
Dim Dy(): Dy = SelDy(A.Dy, ColIxy)
DupRecRxyzFF = DupRecRxy(Dy)
End Function

Private Sub DupRecRxyzColIx__Tst()
Dim Dy(), ColIx&, Act&(), Ept&()
GoSub T0
Exit Sub
T0:
    ColIx = 0
    Dy = Array(Array(1, 2, 3, 4), Array(1, 2, 3), Array(2, 4, 3))
    Ept = LngAy(0, 1)
    GoTo Tst
Tst:
    Act = DupRecRxyzColIx(Dy, ColIx)
    If Not IsEqAy(Act, Ept) Then Stop
    C
    Return
End Sub
