Attribute VB_Name = "MxVbDtaDicCnt"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Dic"
Const CMod$ = CLib & "MxVbDtaDicCnt."

Sub FmtCntDi__Tst()
D FmtCntDi(CntDi(Array(1, 2, 3, 1, 2, 4, "A")))
End Sub
Function FmtCntDi(CntDi As Dictionary, Optional Tit$ = "CntDi", Optional Kn$ = "Key") As String()
Dim Ky$(): Ky = QSrt(SKy(CntDi))
Dim OKy$():    OKy = AmAli(AddSy(UL(Kn), Ky))
Dim OIx$():    OIx = IxCol(Si(Ky), 1)

Dim OCnt$()
    PushUL OCnt, "Cnt"
    Dim K: For Each K In Itr(Ky)
        PushI OCnt, CntDi(K)
    Next
    OCnt = AmAliR(OCnt)

PushIAy FmtCntDi, Box(Tit)
PushIAy FmtCntDi, AddAliStrColAp(OIx, OKy, OCnt)
End Function

Function CntDiwDup(CntDi As Dictionary) As Dictionary
Set CntDiwDup = New Dictionary
Dim Cnt&, K
For Each K In CntDi.Keys
    Cnt = CntDi(K)
    If Cnt > 1 Then CntDiwDup.Add K, Cnt
Next
End Function

Function UniqAyzCntDi(CntDi As Dictionary) As String()
Dim K: For Each K In CntDi.Keys
    If CntDi(K) = 0 Then PushI UniqAyzCntDi, K
Next
End Function

Function CntDi(Ay, Optional C As VbCompareMethod = vbTextCompare) As Dictionary
':CntDi: #Cnt-Dic ! Key-Is-Str & Val-Is-CLng
Dim O As New Dictionary, I
O.CompareMode = C
For Each I In Itr(Ay)
    If O.Exists(I) Then
        O(I) = O(I) + 1
    Else
        O.Add I, CLng(1)
    End If
Next
Set CntDi = O
End Function

Function CntDizDrs(A As Drs, C$) As Dictionary
Set CntDizDrs = CntDi(ColzDrs(A, C))
End Function
