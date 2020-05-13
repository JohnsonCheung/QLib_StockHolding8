Attribute VB_Name = "MxVbAyAe"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Ay.Op"
Const CMod$ = CLib & "MxVbAyAe."

Private Sub AeAtCnt__Tst()
Dim Ay(), At&, Cnt&
GoSub YY
Exit Sub
YY:
    Ay = Array(1, 2, 3, 4, 5)
    At = 1
    Cnt = 2
    Ept = Array(1, 4, 5)
    GoTo Tst
Tst:
    Act = AeAtCnt(Ay, 1, 2)
    C
    Return
End Sub

Function AeAtCnt(Ay, Optional At = 0, Optional Cnt = 1)
Const CSub$ = CMod & "AeAtCnt"
If Cnt <= 0 Then Thw CSub, "Cnt cannot <=0", "At Cnt Ay", At, Cnt, Ay
If Si(Ay) = 0 Then AeAtCnt = Ay: Exit Function
Dim U&: U = UB(Ay)
ChkBet CSub, At, 0, U
Dim NewU&: NewU = U - Cnt
    If At = U - Cnt + 1 Then
        AeAtCnt = ResiAy(Ay, NewU)
        Exit Function
    End If
Dim O: O = Ay
Dim J&: For J = At To U - Cnt
    Asg O(J + Cnt), O(J)
Next
AeAtCnt = ResiAy(Ay, NewU)
End Function

Function AeBlnk(Ay) As String()
Dim I: For Each I In Itr(Ay)
    If Trim(I) <> "" Then PushI AeBlnk, I
Next
End Function

Function AeBlnkAtEnd(A$()) As String()
If Si(A) = 0 Then Exit Function
If LasEle(A) <> "" Then AeBlnkAtEnd = A: Exit Function
Dim J%
For J = UB(A) To 0 Step -1
    If Trim(A(J)) <> "" Then
        Dim O$()
        O = A
        ReDim Preserve O(J)
        AeBlnkAtEnd = O
        Exit Function
    End If
Next
End Function

Function AeBlnkStr(Ay) As String()
Dim X
For Each X In Itr(Ay)
    If Trim(X) <> "" Then
        PushI AeBlnkStr, X
    End If
Next
End Function

Function AeEle(Ay, Ele) 'Rmv Fst-Ele eq to Ele from Ay
AeEle = Ay
Erase AeEle
Dim I: For Each I In Itr(Ay)
    If I <> Ele Then PushI AeEle, I
Next
End Function

Function AeEleAt(Ay, Optional At& = 0, Optional Cnt& = 1)
AeEleAt = AeAtCnt(Ay, At, Cnt)
End Function

Function AeEleLik(Ay, Lik$) As String()
If Si(Ay) = 0 Then Exit Function
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) Like Lik Then AeEleLik = AeEleAt(Ay, J): Exit Function
Next
End Function

Function AeEmpEle(Ay)
Dim O: O = NwAy(Ay)
If Si(Ay) > 0 Then
    Dim X
    For Each X In Itr(Ay)
        PushNonEmp O, X
    Next
End If
AeEmpEle = O
End Function

Function AeEmpEleAtEnd(Ay)
Dim LasU&, U&
Dim O: O = Ay
For LasU = UB(Ay) To 0 Step -1
    If Not IsEmp(O(LasU)) Then
        Exit For
    End If
Next
If LasU = -1 Then
    Erase O
Else
    ReDim Preserve O(LasU)
End If
AeEmpEleAtEnd = O
End Function

Function AeBei(Ay, B As Bei)
With B
    AeBei = AeBE(Ay, .Bix, .Eix)
End With
End Function

Function AeBE(Ay, Bix, Eix)
Const CSub$ = CMod & "AeBE"
Dim U&
U = UB(Ay)
If 0 > Bix Or Bix > U Then Thw CSub, "[Bix] is out of range", "U Bix Eix Ay", UB(Ay), Bix, Eix, Ay
If Bix > Eix Or Eix > U Then Thw CSub, "[Eix] is out of range", "U Bix Eix Ay", UB(Ay), Bix, Eix, Ay
Dim O
    O = Ay
    Dim I&, J&
    I = 0
    For J = Eix + 1 To U
        O(Bix + I) = O(J)
        I = I + 1
    Next
    Dim Cnt&
    Cnt = Eix - Bix + 1
    ReDim Preserve O(U - Cnt)
AeBE = O
End Function

Function AeFstEle(Ay)
AeFstEle = AeEleAt(Ay)
End Function

Function AeFstLas(Ay)
Dim J&
AeFstLas = Ay
Erase AeFstLas
For J = 1 To UB(Ay) - 1
    PushI AeFstLas, Ay(J)
Next
End Function

Function AeFstNEle(Ay, Optional N& = 1)
Dim O: O = NwAy(Ay)
Dim J&
For J = N To UB(Ay)
    Push O, Ay(J)
Next
AeFstNEle = O
End Function

Function AeIxSet(Ay, IxSet As Dictionary)
Dim O: O = Ay: Erase O
Dim J&: For J = 0 To UBound(Ay)
    If Not IxSet.Exists(J) Then PushI O, Ay(J)
Next
AeIxSet = O
End Function

Function AeIxy(Ay, IxyzSrt)
'Fm IxyzSrt : holds index if Ay to be remove.  It has been sorted else will be stop
Ass IsArray(Ay)
Ass IsSrtd(IxyzSrt)
Dim J&
Dim O: O = Ay
For J = UB(IxyzSrt) To 0 Step -1
    O = AeEleAt(O, CLng(IxyzSrt(J)))
Next
AeIxy = O
End Function

Function AeKss(Ay, ExlKss) As String()
Dim O: O = Ay
Dim Lik: For Each Lik In SyzSS(ExlKss)
    O = AeLik(O, Lik)
Next
AeKss = O
End Function

Function AeLasEle(Ay)
AeLasEle = AeEleAt(Ay, UB(Ay))
End Function

Function AeLasNEle(Ay, Optional NEle% = 1)
If NEle = 0 Then AeLasNEle = Ay: Exit Function
Dim O: O = Ay
Select Case Si(Ay)
Case Is > NEle:    ReDim Preserve O(UB(Ay) - NEle)
Case NEle: Erase O
Case Else: Stop
End Select
AeLasNEle = O
End Function

Function AeLik(Ay, Lik) As String()
Dim I: For Each I In Itr(Ay)
    If Not I Like Lik Then PushI AeLik, I
Next
End Function

Function AeNegative(Ay)
Dim I
AeNegative = NwAy(Ay)
For Each I In Itr(Ay)
    If I >= 0 Then
        PushI AeNegative, I
    End If
Next
End Function

Function AeNEle(Ay, Ele, Cnt%)
If Cnt <= 0 Then Stop
AeNEle = NwAy(Ay)
Dim X, C%
C = Cnt
For Each X In Itr(Ay)
    If C = 0 Then
        PushI AeNEle, X
    Else
        If X <> Ele Then
            Push AeNEle, X
        Else
            C = C - 1
        End If
    End If
Next
X:
End Function

Function AePfx(Ay, Pfx$) As String()
Dim I: For Each I In Itr(Ay)
    If Not HasPfx(I, Pfx) Then PushI AePfx, I
Next
End Function
Function AeSfx(Ay, Sfx$) As String()
Dim I: For Each I In Itr(Ay)
    If Not HasSfx(I, Sfx) Then PushI AeSfx, I
Next
End Function

Function AeSngQRmk(Ay) As String()
Dim I, S$
For Each I In Itr(Ay)
    S = I
    If Not IsSngQRmk(S) Then PushI AeSngQRmk, S
Next
End Function


Function AeEmp(Ay)
AeEmp = Ay: Erase AeEmp
Dim I: For Each I In Ay
    If Not IsEmpty(I) Then
        PushI AeEmp, I
    End If
Next
End Function


Function AeKssAy(Ay, KssAy$()) As String()
Dim O: O = Ay
Dim Kss: For Each Kss In KssAy
    O = AeKss(O, Kss)
Next
AeKssAy = O
End Function

Function AeLikAy(Ay, LikAy$()) As String()
Dim O: O = Ay
Dim Lik: For Each Lik In LikAy
    O = AeLik(O, Lik)
Next
AeLikAy = O
End Function

Function AeOneTermLn(Ay) As String()
Stop '
'AeOneTermLn = AePred(Sy, PredzIsOneTermLn)
End Function

Function AeT1Sy(Ay, ExlT1Sy$()) As String()
'Exclude those Ln in Array-Ay its T1 in ExlAmT10
If Si(ExlT1Sy) = 0 Then AeT1Sy = Sy: Exit Function
Stop '
'AeT1Sy = AePred(Sy, PredzInT1Sy(ExlT1Sy))
End Function


Private Sub AeEmpEleAtEnd__Tst()
Dim Ay: Ay = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AeEmpEleAtEnd(Ay)
Ass Si(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub AeEmpEleAtEnd1__Tst()
Dim Ay: Ay = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AeEmpEleAtEnd(Ay)
Ass Si(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub AeBei__Tst()
Dim Ay
Dim Bei1 As Bei
Dim Act
Ay = SplitSpc("a b c d e")
Bei1 = Bei(1, 2)
Act = AeBei(Ay, Bei1)
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub AeBei1__Tst()
Dim Ay
Dim Act
Ay = SplitSpc("a b c d e")
Act = AeBei(Ay, Bei(1, 2))
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub AeIxy__Tst()
Dim Ay(), Ixy
Ay = Array("a", "b", "c", "d", "e", "f")
Ixy = Array(1, 3)
Ept = Array("a", "c", "e", "f")
GoSub Tst
Exit Sub
Tst:
    Act = AeIxy(Ay, Ixy)
    C
    Return
End Sub

Sub AeKss__Tst()
Dim Sy$(), Kss$
GoSub Z
GoSub T0
Exit Sub
T0:
    Sy = SyzSS("A B C CD E E1 E3")
    Kss = "C* E*"
    Ept = SyzSS("A B")
    GoTo Tst
Z:
    D AeKss(SyzSS("A B C CD E E1 E3"), "C* E*")
    Return
Tst:
    Act = AeKss(Sy, Kss)
    C
    Return
End Sub

Function AeAt(Ay, At&)
Dim O: O = Ay
Dim U&: U = UB(Ay)
Dim J&: For J = At To U - 1
    O(J) = O(J + 1)
Next
ReDim Preserve O(U - 1)
AeAt = O
End Function
