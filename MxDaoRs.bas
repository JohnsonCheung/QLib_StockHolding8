Attribute VB_Name = "MxDaoRs"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDaoRs."
Private Sub CRszSkvap__Tst()
Dim Rs As DAO.Recordset: Set Rs = CRszSkvap("Att", 1, "AAA")
Stop
End Sub
Function CRszSkvap(T, ParamArray Skvap()) As DAO.Recordset
Dim Skvy(): Skvy = Skvap
Set CRszSkvap = RszSkvy(CDb, T, Skvy)
End Function
Function RszSkvap(D As Database, T, ParamArray Skvap()) As DAO.Recordset 'return a Rs in edit-mode of @T if there is @Skvap else in AddNew-mode assuming @T can be insert with @Skvy only
Dim V(): V = Skvap: Set RszSkvap = RszSkvy(D, T, V)
End Function
Function RszSkvy(D As Database, T, Skvy()) As DAO.Recordset 'return a Rs in edit-mode of @T if there is @Skvap else in AddNew-mode assuming @T can be insert with @Skvy only
Dim Q$: Q = SqlSelStar_T_Skvy(D, T, Skvy)
If HasRecQ(D, Q) Then Set RszSkvy = RszQ(D, Q): RszSkvy.Edit: Exit Function
Set RszSkvy = InsRsSkvy(D, T, Skvy)
End Function
Function InsRsSkvy(D As Database, T, Skvy()) As DAO.Recordset '#Insert-Rs-By-Skvy# insert a new rs to @T using @Skvy, keep not Update.
Dim O As DAO.Recordset: Set O = RszT(D, T)
Dim J%
With O
    .AddNew
    While Not .EOF
        Dim F: For Each F In SkFny(D, T)
            .Fields(F).Value = Skvy(J)
            J = J + 1
        Next
    Wend
End With
End Function

Function AetzRs(Rs As DAO.Recordset, Optional F = 0) As Dictionary
Set AetzRs = New Dictionary
With Rs
    While Not .EOF
        PushEle AetzRs, .Fields(F).Value
        .MoveNext
    Wend
End With
End Function

Sub AsgRs(A As DAO.Recordset, ParamArray OAp())
Dim F As DAO.Field, J%, U%
Dim Av(): Av = OAp
U = UB(Av)
For Each F In A.Fields
    OAp(J) = F.Value
    If J = U Then Exit Sub
    J = J + 1
Next
End Sub

Function AvzRsF(A As DAO.Recordset, Optional Fld = 0) As Variant(): AvzRsF = IntozRs(EmpAv, A, Fld): End Function
Sub BrwRs(A As DAO.Recordset): BrwDrs DrszRs(A): End Sub
Sub BrwRec(A As DAO.Recordset): BrwAy FmtRec(A): End Sub
Function CvRs2(A) As DAO.Recordset2: Set CvRs2 = A: End Function
Function CvRs(A) As DAO.Recordset: Set CvRs = A: End Function
Sub DltRs(A As DAO.Recordset)
With A
    While Not .EOF
        .Delete
        .MoveNext
    Wend
End With
End Sub

Sub DmpRec(A As DAO.Recordset, Optional FF$): DmpAy FmtRec(A, FF): End Sub
Function DrszRs(A As DAO.Recordset) As Drs: DrszRs = Drs(FnyzRs(A), DyzRs(A)): End Function
Function DrzRs(A As DAO.Recordset, Optional FF$) As Variant(): DrzRs = DrzFds(A.Fields, FF): End Function
Function DrzRsFny(A As DAO.Recordset, Fny$()) As Variant(): DrzRsFny = DrzFdsFny(A.Fields, Fny): End Function
Function FnyzRs(A As DAO.Recordset) As String(): FnyzRs = Itn(A.Fields): End Function
Function HasRecFxQ(Fx$, Q$): HasRecFxQ = HasReczArs(ArszFxQ(Fx, Q)): End Function
Function HasBlnkzFxwc(Fx$, W$, C$) As Boolean
Dim Wh$: Wh = FldIsBlnkBexp(C)
Dim Q$: Q = SqlSel_F_T(C, AxTbn(W), Wh)
HasBlnkzFxwc = HasRecFxQ(Fx, Q)
End Function

Function HasRec(R As DAO.Recordset) As Boolean: HasRec = Not NoRec(R): End Function
Function HasRecRsFeq(R As DAO.Recordset, F, Eqval) As Boolean
If NoRec(R) Then Exit Function
With R
    .MoveFirst
    While Not .EOF
        If .Fields(F) = Eqval Then HasRecRsFeq = True: Exit Function
        .MoveNext
    Wend
End With
End Function

Sub InsRszDy(A As DAO.Recordset, Dy())
Dim Dr: For Each Dr In Itr(Dy)
    InsRs A, Dr
Next
End Sub

Sub InsRs(Rs As DAO.Recordset, Dr)
Rs.AddNew
SetRs Rs, Dr
Rs.Update
End Sub

Sub InsRszAp(Rs As DAO.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
InsRs Rs, Dr
End Sub

Function NoRec(A As DAO.Recordset) As Boolean
If Not A.EOF Then Exit Function
If Not A.BOF Then Exit Function
NoRec = True
End Function

Function NReczRs&(A As DAO.Recordset)
If NoRec(A) Then Exit Function
Dim O&
With A
    .MoveFirst
    While Not .EOF
        O = O + 1
        .MoveNext
    Wend
    .MoveFirst
End With
NReczRs = O
End Function

Sub RsDlt(A As DAO.Recordset)
With A
    If .EOF Then Exit Sub
    If .BOF Then Exit Sub
    .Delete
End With
End Sub

Function RsLin(A As DAO.Recordset, Optional Sep$ = " ")
RsLin = Join(DrzRs(A), Sep)
End Function

Function JnRsFny(A As DAO.Recordset, Fny$(), Optional Sep$ = " ") As String()

End Function


Sub SetRs(Rs As DAO.Recordset, Dr)
Const CSub$ = CMod & "SetRs"
If Si(Dr) <> Rs.Fields.Count Then
    Thw CSub, "Si of Rs & Dr are diff", _
        "Si-Rs and Si-Dr Rs-Fny Dr", Rs.Fields.Count, Si(Dr), Itn(Rs.Fields), Dr
End If
Dim V, J%
For Each V In Dr
    If IsEmpty(V) Then
        Rs(J).Value = Rs(J).DefaultValue
    Else
        Rs(J).Value = V
    End If
    J = J + 1
Next
End Sub

Sub UpdRsV(Rs As DAO.Recordset, V)
Rs.Edit
Rs.Fields(0).Value = V
Rs.Update
End Sub

Sub UpdRs(Rs As DAO.Recordset, Dr)
Rs.Edit
SetRs Rs, Dr
Rs.Update
End Sub

Sub UpdRszAp(Rs As DAO.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
UpdRs Rs, Dr
End Sub

'**PimyzRs
Function IntAyzRs(A As DAO.Recordset, Optional Fld = 0) As Integer(): IntAyzRs = IntozRs(IntAyzRs, A, Fld): End Function
Function LngAyzRs(A As DAO.Recordset, Optional Fld = 0) As Long(): LngAyzRs = IntozRs(LngAyzRs, A, Fld): End Function
Function SyzRs(A As DAO.Recordset, Optional F = 0) As String(): SyzRs = IntozRs(EmpSy, A, F): End Function
Private Function IntozRs(Into, Rs As Recordset, Optional Fld = 0):
IntozRs = NwAy(Into)
While Not Rs.EOF
    PushI IntozRs, Nz(Rs(Fld).Value, Empty)
    Rs.MoveNext
Wend
End Function

Function SqzRs(A As DAO.Recordset, Optional IsIncFldn As Boolean) As Variant(): SqzRs = SqzDy(DyzRs(A, IsIncFldn)): End Function
