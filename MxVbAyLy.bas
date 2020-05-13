Attribute VB_Name = "MxVbAyLy"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbAyLy."


Private Sub IndLy__Tst()
Dim IndtSrc$(), K$
GoSub Z
GoSub T0
Exit Sub
T0:
    K = "A"
    Erase XX
    X "A Bc"
    X " 1"
    X " 2"
    X "A 2"
    IndtSrc = XX
    Erase XX
    Ept = Sy("1", "2")
    GoTo Tst
Tst:
    Act = IndLy(IndtSrc, K)
    C
    Return
Z:
    K = "A"
    Erase XX
    X "A Bc"
    X " 1"
    X " 2"
    X "A Bc"
    X " 1 2"
    X " 2 3"
    IndtSrc = XX
    Erase XX
    D IndLy(IndtSrc, K)
    Return
End Sub

Function IndLy(IndtSrc$(), Key$) As String()
Dim O$()
Dim L, Fnd As Boolean, IsNewSection As Boolean, IsFstChrSpc As Boolean, FstA%, Hit As Boolean
Const SpcAsc% = 32
For Each L In Itr(IndtSrc)
    If Fst2Chr(LTrim(L)) = "--" Then GoTo Nxt
    FstA = FstAsc(L)
    IsNewSection = IsAscUCas(FstA)
    If IsNewSection Then
        Hit = T1(L) = Key
    End If
    
    IsFstChrSpc = FstA = SpcAsc
    Select Case True
    Case IsNewSection And Not Fnd And Hit: Fnd = True
    Case IsNewSection And Fnd:             IndLy = O: Exit Function
    Case Fnd And IsFstChrSpc:              PushI O, Trim(L)
    End Select
Nxt:
Next
If Fnd Then IndLy = O: Exit Function
End Function
