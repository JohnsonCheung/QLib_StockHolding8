Attribute VB_Name = "MxVbRmk"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVrmk."
'--IsVrmkLn
Function IsVrmkLn(L) As Boolean: IsVrmkLn = FstChr(LTrim(L)) = "'": End Function
Function IsRmkOrBlnk(Ln) As Boolean
Select Case True
Case IsBlnk(Ln), IsRmkln(Ln): IsRmkOrBlnk = True: Exit Function
End Select
End Function

Private Sub FstVrmk__Tst(): BrwAy FstVrmk(SrcM): End Sub
Function FstVrmk(Src$()) As String() ':Ly #First-VbRmk#
Dim B&: B = VrmkBix(Src): If B < 0 Then Exit Function
Dim E&: E = VrmkEix(Src, B)
FstVrmk = AwBE(Src, B, E)
End Function
Function FstVrmkl$(Src$()): FstVrmkl = JnCrLf(FstVrmk(Src)): End Function

'--BrkVrmk
Function BrkVrmk(Ln) As S12
Dim P%: P = VrmkPos(Ln): If P = 0 Then BrkVrmk = S12(Trim(Ln), ""): Exit Function
BrkVrmk = S12(Left(Ln, P - 1), Mid(Ln, P + 1))
End Function
Private Function VrmkBix&(Src$(), Optional FmIx& = 0)
Dim J&: For J = FmIx To UB(Src)
    If IsRmkln(Src(J)) Then VrmkBix = J: Exit Function
Next
VrmkBix = -1
End Function
Private Function VrmkEix&(Src$(), Bix&)
Dim J&: For J = Bix To UB(Src)
    If Not IsRmkln(Src(J)) Then VrmkEix = J - 1: Exit Function
Next
VrmkEix = UB(Src)
End Function

'--RmvVrmk
Private Sub RmvVrmkAndVstr__Tst()
TimFun "W1Tst1"
TimFun "W1Tst2" ' Quicker
End Sub
Function RmvVrmk(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushNB RmvVrmk, RmvVrmkzL(L)
Next
End Function
Function RmvVrmkzL$(Ln): RmvVrmkzL = LeftOrAll(Ln, VrmkPos(Ln)): End Function
Function RmvVrmkAndVstr(Ly$()) As String(): RmvVrmkAndVstr = W1V2(Ly): End Function
Private Function W1V1(Ly$()) As String(): W1V1 = RmvVrmk(RmvVstrzS(Ly)): End Function
Private Function W1V2(Ly$()) As String()
Dim L: For Each L In Itr(Ly)
    PushI W1V2, RmvVrmkzL(RmvVstr(L))
Next
End Function
Private Sub W1Tst1(): W1V1 SrcP: End Sub
Private Sub W1Tst2(): W1V2 SrcP: End Sub

'--VrmkPos
Private Sub VrmkPos__Tst()
Dim I, O$(), L$, P%
For Each I In AwSubStr(AwSubStr(SrczP(CPj), "'"), """")
    P = VrmkPos(I)
    If P = 0 Then
        PushI O, I
    Else
        PushI O, I & vbCrLf & Dup(" ", P - 1) & "^"
    End If
Next
Vc O
End Sub
Function VrmkPos%(Ln)
Dim Py%(): Py = SubStrPosy(Ln, "'"): If Si(Py) = 0 Then Exit Function
Dim P: For Each P In Py
    If IsVrmkPos(Ln, P) Then VrmkPos = P: Exit Function
Next
End Function
Function IsVrmkPos(Ln, SngQPos) As Boolean: IsVrmkPos = Not W1IsInDblQPos(Ln, SngQPos): End Function
Private Function W1IsInDblQPos(Ln, Pos) As Boolean
Dim Bef$: Bef = Left(Ln, Pos - 1)
W1IsInDblQPos = IsOdd(SubStrCnt(Bef, vbDblQ))
End Function
