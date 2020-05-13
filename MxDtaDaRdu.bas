Attribute VB_Name = "MxDtaDaRdu"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaRdu."
Private Type RduDrs  ' #Reduced-Drs ! if a drs col all val are sam, mov those cols to @RduColDic (Dic-of-coln-to-val).
    Drs As Drs       '              ! the drs aft rmv the sam val col
    CnstCol As Dictionary '        ! one entry is one col.  Key is colNm and val is colNm val.
    FldSum As Dictionary ' For all Numerice column, sum them
    DupCol As Dictionary ' Key=DupCol Val=ColWillShw
End Type

Function FmtDrsR(D As Drs, Optional DrsFmts$, Optional Nm$ = "Drs1") As String()  ' format drs @D reducely with format option
FmtDrsR = FmtDrsRO(D, DrsFmtozS(DrsFmts), Nm)
End Function

Function FmtDrs(D As Drs, Optional Nm$ = "Drs1") As String() ' format drs @D reducely with defualt format option
FmtDrs = FmtDrsR(D, Nm)
End Function
'---===============================================
Private Sub FmtDrsR__Tst()
Brw FmtDrsR(SampDrs)
End Sub

Function FmtDrsRO(D As Drs, Opt As DrsFmto, Optional Nm$ = "Drs1") As String() ' format drs @D reducely with defualt format option
Dim R As RduDrs: R = W1Rdu(D)
Dim Ly1$(): Ly1 = RmvLasEle(FmtDic(R.DupCol, H12:="Skip See"))
Dim Ly2$(): Ly2 = RmvLasEle(FmtDic(R.CnstCol, H12:="Cnst Val"))
Dim Ly3$(): Ly3 = RmvLasEle(FmtDic(R.FldSum, H12:="Col Sum"))
PushIAy FmtDrsRO, FmtStrColAp(Ly1, Ly2, Ly3)
PushIAy FmtDrsRO, FmtDrsNO(R.Drs, Opt)
End Function

Private Function W1Rdu(D As Drs) As RduDrs
If NoReczDrs(D) Then W1Rdu = W1Emp(D): Exit Function
Set W1Rdu.CnstCol = W1Cnst(D)
Set W1Rdu.DupCol = W1Dup(D)
Set W1Rdu.FldSum = W1Sum(D)
Dim Fny$(): Fny = MgeSKy(W1Rdu.CnstCol, W1Rdu.DupCol)
W1Rdu.Drs = DrpColzDrsFny(D, Fny)
End Function

Private Function W1Sum(A As Drs) As Dictionary
':FldSum: :DiNumFqSum '#Fld-Sum-Di# Key is NumFldNm and Val is the sum of the fld in Dbl
Dim O As New Dictionary, Sum
Dim F: For Each F In A.Fny
    Sum = W1SumCol(A, F)
    If Not IsEmpty(Sum) Then
        O.Add F, Sum
    End If
Next
Set W1Sum = O
End Function
Private Function W1SumCol(A As Drs, C)
If C = "Ix" Then Exit Function
Dim O#
Dim Ix%: Ix = EleIx(A.Fny, C)
Dim V, Dr: For Each Dr In Itr(A.Dy)
    If UB(Dr) >= Ix Then        ' Dr may have less field the column-@C, which is convert to *Ix
        V = Dr(Ix)
        If Not IsEmpty(V) Then
            If Not IsNum(V) Then Exit Function
        End If
        O = O + V
    End If
Next
W1SumCol = O
End Function

Private Function W1Emp(D As Drs) As RduDrs
With W1Emp
       .Drs = D
Set .CnstCol = New Dictionary
Set .DupCol = New Dictionary
Set .FldSum = New Dictionary
End With
End Function

Private Function W1Cnst(A As Drs) As Dictionary
':CnstColDi: :Di ! Key is fld nm of @A with all rec has same value.  Val is that column value
Dim NCol%: NCol = NColzDy(A.Dy)
Dim Dy(), Fny$()
Fny = A.Fny
Dy = A.Dy
Dim O As New Dictionary
Dim J%: For J = 0 To NCol - 1
    If IsEqzAllEle(ColzDy(Dy, J)) Then
        O.Add Fny(J), Dy(0)(J)
    End If
Next
Set W1Cnst = O
End Function

Private Sub W1Dup__Tst()
BrwDic W1Dup(MdDrsP)
End Sub

Private Function W1Dup(A As Drs) As Dictionary ' duplicated column dictionary
Set W1Dup = New Dictionary
If NoReczDrs(A) Then Exit Function
Dim Fny$(): Fny = A.Fny
Dim DoneCix%()
Dim Dy(): Dy = A.Dy
Dim U%: U = UB(Fny)
Dim J%: For J = 0 To U - 1
    If HasEle(DoneCix, J) Then GoTo Nxt
    Dim I%: For I = J + 1 To U
        If IsEqCol(Dy, I, J) Then
            W1Dup.Add Fny(I), Fny(J)
            PushI DoneCix, I
        End If
    Next
Nxt:
Next
End Function

'---===================================
Private Sub DmpDrsRO__Tst()
Dim Drs As Drs
GoSub Z
Exit Sub
Z:
    Dim AF1$, AF2$, BF1$, BF2$
    AF1 = RplVBar("AA|BBBB|CCCC")
    AF2 = RplVBar("AA|BBBB|CCCC")
    BF1 = RplVBar("AA|BBBB|CCCC")
    BF2 = RplVBar("AA|BBBB|CCCC")
    Dim Dr1(): Dr1 = Array(AF1, AF2)
    Dim Dr2(): Dr2 = Array(BF1, BF2)
    Drs = DrszFF("A B", CvAv(Array(Dr1, Dr2)))
    DmpDrsR Drs
    Return
End Sub

Sub DmpDrs(D As Drs)
DmpDrsR D
End Sub
Sub DmpDrsN(D As Drs)
DmpAy FmtDrsN(D)
End Sub
Sub DmpDrsR(A As Drs)
DmpAy FmtDrsR(A)
End Sub
