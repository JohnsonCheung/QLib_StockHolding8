Attribute VB_Name = "MxDaoDbDta"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CNs$ = "Db"
Const CMod$ = CLib & "MxDaoDbDta."
Type StrColPair: C1() As String: C2() As String: End Type
Function LngAyzQ(D As Database, Q) As Long(): LngAyzQ = LngAyzRs(Rs(D, Q)): End Function
Function SyzCQ(Q) As String(): SyzCQ = SyzQ(CDb, Q): End Function
Function SyzQ(D As Database, Q) As String(): SyzQ = SyzRs(Rs(D, Q)): End Function

Private Sub Rs__Tst()
Shell "Subst N: c:\subst\users\user\desktop", vbHide
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
BrwAy CsyzRs(Rs(TmpDb, S))
End Sub
Function StrColzTF(D As Database, T, F, Optional Bexp$) As String():         StrColzTF = StrColzRs(RszT(D, T), F):                    End Function
Function StrColzQ(D As Database, Q, Optional Bexp$) As String():              StrColzQ = StrColzRs(Rs(D, Q)):                         End Function
Sub AsgStrColPair(P As StrColPair, OC1$(), OC2$()): OC1 = P.C1: OC2 = P.C2: End Sub
Function StrColPair(D As Database, T, F12$, Optional Bexp$) As StrColPair: StrColPair = StrColPairzQ(D, SqlSel_F12_Fm(F12, T, Bexp)): End Function
Function StrColPairzQ(D As Database, Q) As StrColPair:                    StrColPairzQ = StrColPairzRs(Rs(D, Q)):                     End Function
Function StrColPairzRs(R As DAO.Recordset) As StrColPair
Dim O As StrColPair
With R
    If Not .EOF Then .MoveFirst
    While Not .EOF
        PushI O.C1, Nz(.Fields(0).Value, "")
        PushI O.C2, Nz(.Fields(1).Value, "")
        .MoveNext
    Wend
End With
End Function

Function ColzRs(R As DAO.Recordset, Optional F = 0) As Variant(): ColzRs = IntoColzRs(EmpAv, R, F): End Function
Function StrColzRs(R As DAO.Recordset, Optional F = 0) As String(): StrColzRs = IntoColzRs(EmpSy, R, F): End Function
Function IntAyzQ(D As Database, Q) As Integer(): IntAyzQ = IntoColzRs(EmpIntAy, Rs(D, Q)): End Function
Function SyzF(D As Database, T, F$) As String(): SyzF = SyzRs(RszF(D, T, F)): End Function
Function SyzTF(D As Database, TF$) As String(): SyzTF = SyzRs(RszTF(D, TF)): End Function
Private Function IntozF(Into, D As Database, T, F$): IntozF = IntoColzRs(Into, RszF(D, T, F)): End Function
Private Function IntoColzRs(IntoAy, R As DAO.Recordset, Optional F = 0)
Dim O: O = IntoAy: Erase O
With R
    If Not .EOF Then .MoveFirst
    While Not .EOF
        PushI O, .Fields(F).Value
        .MoveNext
    Wend
End With
IntoColzRs = O
End Function
Function FnyzQ(D As Database, Q) As String(): FnyzQ = FnyzRs(Rs(D, Q)): End Function
Private Sub FnyzQ__Tst()
Dim Db As Database
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
DmpAy FnyzQ(Db, S)
End Sub
Function DrzQ(D As Database, Q) As Variant(): DrzQ = DrzRs(Rs(D, Q)): End Function
Function DyzQ(D As Database, Q) As Variant(): DyzQ = DyzRs(Rs(D, Q)): End Function
Function DyzRs(A As DAO.Recordset, Optional IsIncFldn As Boolean) As Variant()
If IsIncFldn Then
    PushI DyzRs, FnyzRs(A)
End If
If Not HasRec(A) Then Exit Function
With A
    .MoveFirst
    While Not .EOF
        PushI DyzRs, DrzFds(.Fields)
        .MoveNext
    Wend
    .MoveFirst
End With
End Function

