Attribute VB_Name = "MxVbAyQSrt"
Option Explicit
Option Compare Text
Const CNs$ = "Ay.Srt"
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxVbAyQSrt."
Enum eOrd: eByAsc: eByDes: End Enum

Private Sub QSrt__Tst()
Dim Ay, Ord As eOrd
GoSub T0
GoSub T1
Exit Sub
T0:
    Ay = Array(1, 2, 3, 4, 0, 1, 1, 5)
    Ord = eByAsc
    Ept = Array(0, 1, 1, 1, 2, 3, 4, 5)
    GoTo Tst
T1:
    Ay = Array(1, 2, 4, 87, 4, 2)
    Ord = eByDes
    Ept = Array(87, 4, 4, 2, 2, 1)
    GoTo Tst
Tst:
    Act = QSrt(Ay, Ord)
    C
    Return
End Sub

Function QSrt(Ay, Optional By As eOrd)
If Si(Ay) = 0 Then QSrt = Ay: Exit Function
Dim O: O = W1Srt(Ay)
If By = eByDes Then O = RevIAy(O)
QSrt = O
End Function

Private Function W1Srt(Ay)
Select Case Si(Ay)
Case 0, 1: W1Srt = Ay
Case 2 And Ay(0) <= Ay(1)
    W1Srt = Ay
Case 2
    Dim O: O = NwAy(Ay)
    PushI O, Ay(1)
    PushI O, Ay(0)
    W1Srt = O
Case Else
    Dim P: P = LasEle(Ay)
    Dim X: X = RmvLasEle(Ay)
    Dim A: A = W1Srt(AwLE(X, P))
    Dim B: B = W1Srt(AwGT(X, P))
    W1Srt = FlatAp(A, P, B)
End Select
End Function
