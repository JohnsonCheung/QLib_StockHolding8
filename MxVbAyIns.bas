Attribute VB_Name = "MxVbAyIns"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay.Op"
Const CMod$ = CLib & "MxVbAyIns."

Function Ins2Ele(Ay, E1, E2, Optional Bef = 0)
Ins2Ele = InsAy(Ay, Array(E1, E2), Bef)
End Function

Private Sub InsBef__Tst()
Dim Ay, M, Bef&
'
Ay = Array(1, 2, 3)
M = "X"
Ept = Array("X", 1, 2, 3)
GoSub Tst
'
Exit Sub
Tst:
    Act = InsBef(Ay, M, Bef)
    C
    Return
End Sub

Function InsBef(Ay, Optional Ele, Optional Bef = 0)
InsBef = InsAy(Ay, Array(Ele), Bef)
End Function

Function InsAft(Ay, Optional Ele, Optional Aft = 0)
InsAft = InsAy(Ay, Array(Ele), Aft + 1)
End Function
Private Sub InsAy__Tst()
Dim AyA, AyB, Bef&
GoSub YY
Exit Sub
YY:
    AyA = Array(3)
    AyB = Array(1)
    Bef = 0
    Ept = Array(1, 3)
    GoTo Tst
Tst:
    Act = InsAy(AyA, AyB, Bef)
    C
    Return
End Sub
Function InsAy(AyA, AyB, Optional Bef = 0)
Dim O, BSi&
BSi = Si(AyB)
O = InsEle(AyA, Bef, BSi)
Dim J&: For J = 0 To BSi - 1
    O(Bef + J) = AyB(J)
Next
InsAy = O
End Function

Private Sub ReBase__Tst()
Dim Ay&(): ReDim Ay(9)
Dim J%: For J = 0 To 9: Ay(J) = J: Next
Dim Act: Act = ReBase(Ay, 100)
Debug.Assert LBound(Act) = 100
Debug.Assert UBound(Act) = 109
For J = 100 To 109
    Debug.Assert Act(J) = J - 100
Next
End Sub

Function ReBase(Ay, LBix)
'@Ay : assume Ay-LBound is 0
'Ret  : new re-based ay of LBound is @LBIx preserve.  Note standard redim preserve X(F To T) does not work  @@
Dim UBix&: UBix = LBix + UB(Ay)
Dim O:        O = Ay: ReDim O(LBix To UBix)
Dim J&:       J = LBix
Dim V: For Each V In Ay
    O(J) = V
    J = J + 1
Next
ReBase = O
End Function
Function ResiAy(Ay, U&)
'Ret : new ay redim preserve @Ay to @U
Dim O: O = Ay
If U < 0 Then
    Erase O
Else
    ReDim Preserve O(U)
End If
ResiAy = O
End Function

Function InsEle(Ay, Bef, Optional Cnt = 1, Optional Ele)
Dim U&: U = UB(Ay)
Dim NewU&: NewU = U + Cnt
Dim O: O = Ay: ReDim Preserve O(NewU)

Dim F&, T&
Dim X&: X = NewU + Bef
Dim J&: For J = Bef To U
    T = X - J
    F = T - Cnt
    O(T) = Ay(F)
    Ay(F) = Ele
Next
InsEle = O
End Function

Private Sub InsEle__Tst()
Dim Ay(), At&, Cnt&, Ele
GoSub T
Exit Sub
T:
    Ay = Array(1, 2, 3)
    At = 1
    Cnt = 3
    Ele = Empty
    Ept = Array(1, Empty, Empty, Empty, 2, 3)
    GoTo Tst
Tst:
    Act = InsEle(Ay, At, Cnt)
    C
    Return
End Sub
