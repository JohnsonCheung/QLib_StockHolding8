Attribute VB_Name = "MxVbDtaIxBei"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Dta"
Const CMod$ = CLib & "MxVbDtaIxBei."
Type FCnt
    FmLno As Long
    Cnt As Long
End Type
Type Bei: Bix As Long: Eix As Long: End Type
Function BeiUB&(A() As Bei): BeiUB = BeiSi(A) - 1: End Function
Function BeiSi&(A() As Bei): On Error Resume Next: BeiSi = UBound(A) + 1: End Function
Function FCnt(FmLno, Cnt) As FCnt
If FmLno <= 0 Then Exit Function
If Cnt <= 0 Then Exit Function
FCnt.FmLno = FmLno
FCnt.Cnt = Cnt
End Function
Function FCntzBei(A As Bei) As FCnt
With A
    FCntzBei = FCnt(.Bix + 1, LinCntzBei(A))
End With
End Function

Sub BrwBeiAy(A() As Bei)
BrwStr BeiAyStr(A)
End Sub

Function BeiStr$(A As Bei)
BeiStr = FmtQQ("Bei ? ?", A.Bix, A.Eix)
End Function

Function BeiAyStr$(A() As Bei)
Dim O$()
Dim J&: For J = 0 To BeiUB(A)
    With A(J)
        PushI O, FmtQQ("?, ?", .Bix, .Eix)
    End With
Next
BeiAyStr = FmtQQ("BeiAy(?)", JnCommaSpc(O))
End Function

Function BeiyzBooly(Booly() As Boolean) As Bei()
Dim U&: U = UB(Booly): If U = -1 Then Exit Function
Dim B&(): B = W1Bixy(Booly)
Dim J&: For J = 0 To UB(B)
    PushBei BeiyzBooly, Bei(B(J), W1Eix(Booly, B(J)))
Next
End Function
Private Function W1Bixy(Booly() As Boolean) As Long()
Dim Las As Boolean, Fst As Boolean
Dim J&: For J = 0 To UB(Booly)
    Select Case True
    Case (Fst Or Not Las) And Booly(J): Fst = False: PushI W1Bixy, J
    End Select
    If Fst Then Fst = False
    Las = Booly(J)
Next
End Function
Private Function W1Eix&(Booly() As Boolean, Bix&)
Dim U&: U = UB(Booly)
Dim J&: For J = Bix + 1 To U
    If Not Booly(J) Then W1Eix = J - 1: Exit Function
Next
W1Eix = U
End Function

Function Bei(Bix, Eix) As Bei
Select Case True
Case 0 > Bix, -1 > Eix, Bix > Eix: Bei = EmpBei
Case Else
    Bei.Bix = Bix
    Bei.Eix = Eix
End Select
End Function

Function EmpBei() As Bei
EmpBei.Bix = 0
EmpBei.Eix = -1
End Function

Sub PushBei(O() As Bei, M As Bei)
Dim N&: N = BeiSi(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Sub PushBeiAy(O() As Bei, A() As Bei)
Dim J&: For J = 0 To BeiUB(A)
    PushBei O, A(J)
Next
End Sub
Function AddBei(A As Bei, B As Bei) As Bei()
PushBei AddBei, A
PushBei AddBei, B
End Function

Function BetBei(Ix, A As Bei) As Boolean
BetBei = IsBet(Ix, A.Bix, A.Eix - 1)
End Function

Function CntzBei&(A As Bei)
Dim O&
O = A.Eix - A.Bix
If O < 0 Then Stop
CntzBei = O
End Function
Function BeizFC(Bix, Cnt) As Bei
BeizFC = Bei(Bix, Bix + Cnt - 1)
End Function

Function IsEqBeiy(A() As Bei, B() As Bei) As Boolean
Dim U&: U = BeiUB(A)
If U <> BeiUB(B) Then Exit Function
Dim J&: For J = 0 To U
    If Not IsEqBei(A(J), B(J)) Then Exit Function
Next
IsEqBeiy = True
End Function
Function IsEmpBei(A As Bei) As Boolean
Select Case True
Case A.Bix < 0, A.Eix < 0: IsEmpBei = True
End Select
End Function
Function IsBeiInOrd(A() As Bei) As Boolean
Dim J&: For J = 0 To BeiUB(A)
    With FCntzBei(A(J))
        If .FmLno = 0 Then Exit Function
        If .Cnt = 0 Then Exit Function
        If .FmLno + .Cnt > FCntzBei(A(J + 1)).FmLno Then Exit Function
    End With
Next
IsBeiInOrd = True
End Function

Function Positive(N)
If N > 0 Then Positive = N
End Function
Function LinCntzBei&(A As Bei)
LinCntzBei = Positive(A.Eix - A.Bix)
End Function
Function LinCntzBeiAy&(A() As Bei)
Dim J&, O&
For J = 0 To BeiUB(A)
    O = O + LinCntzBei(A(J))
Next
LinCntzBeiAy = O
End Function

Function IsEqBei(A As Bei, B As Bei) As Boolean
With A
    If .Bix <> B.Bix Then Exit Function
    If .Eix <> B.Eix Then Exit Function
End With
IsEqBei = True
End Function
