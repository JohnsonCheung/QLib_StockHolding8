Attribute VB_Name = "MxDtaDaSrtRxyDy"
Option Explicit
Option Compare Text
Const CNs$ = "Srt.Dy"
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxDtaDaSrtRxyDy."
Dim A_Dy()
Dim A_IsDesAy() As Boolean
Private Type LE_GT
    LE() As Long
    GT() As Long
End Type
Private Sub RxyzSrtDy__Tst()
Dim Dy(), IsDesAy() As Boolean
GoSub T0
GoSub T1
Exit Sub
T0:
    Dy = DyzVbl("2 a C|1 c B|3 b A")
    Ept = LngAy(1, 0, 2)
    Erase IsDesAy
    GoTo Tst
T1:
    Dy = DyzVbl("2 a C|1 c B|3 b A")
    Ept = LngAy(2, 0, 1)
    IsDesAy = BoolAy("t..")
    GoTo Tst
Tst:
    Act = RxyzSrtDy(Dy, IsDesAy)
    C
    Return
End Sub

Function RxyzSrtDyByKey(Dy(), K As DySrtKey) As Long()
'Ret :Rxy ! #Sorting-Dy-Row-Index-Ay# It is Rxy @@
':Rxy: :Long() ! #Row-Index-Ay# it is pointing to some Row.  All ele is uniq.  Each ele is between 0 & U.
RxyzSrtDyByKey = RxyzSrtDy(SelDy(Dy, K.Cxy), K.IsDes)
End Function

Function RxyzSrtDy(Dy(), Optional ColDesAy) As Long()
'Ret: :Rxy ! using all column to sort @Dy into @@Rxy
If IsEmpAy(Dy) Then Exit Function
Dim L&(): L = LngSno(Si(Dy), 0)
     A_Dy = Dy
A_IsDesAy = IsDesAy_(ColDesAy, UB(Dy(0)))
RxyzSrtDy = Srt_(L)
Erase A_Dy
Erase A_IsDesAy
End Function
Private Function IsDesAy_(ColDesAy, U&) As Boolean()
If IsBoolAy(ColDesAy) Then
    If UB(ColDesAy) = U Then
        IsDesAy_ = ColDesAy
        Exit Function
    End If
End If
ReDim IsDesAy_(U)
End Function
Private Function LE_GT(Ixy&(), I&) As LE_GT
'Ret : Subset-of-Ixy so that each ele is LE than I
Dim LE&(), GT&()
Dim Dr: Dr = A_Dy(I)
Dim J: For Each J In Ixy
    If IsLE_(J, Dr) Then
        PushI LE, J
    Else
        PushI GT, J
    End If
Next
LE_GT.LE = LE
LE_GT.GT = GT
End Function

Private Function GT_(Ixy&(), I&) As Long()
'Ret : Subset-of-Ixy so that each ele is GT than I
Dim KeyB: KeyB = A_Dy(I)
Dim J: For Each J In Ixy
    If Not IsLE_(J, KeyB) Then PushI GT_, J
Next
End Function

Private Function IsLE_(IxA, Dr) As Boolean
'Ret : true if @A is LE than @B
Dim DrA: DrA = A_Dy(IxA)
IsLE_ = IsLEzAy(DrA, Dr, A_IsDesAy)
End Function

Private Function Srt_(Ixy&()) As Long()
Dim O&()
    Select Case UB(Ixy)
    Case -1
    Case 0: O = Ixy
    Case 1:
        O = Swap_(Ixy)
    Case Else
        Dim C&(): C = Ixy           'Cur
        Dim P&:   P = Pop(C)        'Pivot
        Dim A As LE_GT: A = LE_GT(C, P)
        Dim L&(): L = Srt_(A.LE)      'Low
        Dim H&(): H = Srt_(A.GT)      'High

        PushIAy O, L
          PushI O, P
        PushIAy O, H
    End Select
Srt_ = O
End Function

Private Function Swap_(Ixy2&()) As Long()
Dim KeyB: KeyB = A_Dy(Ixy2(1))
If IsLE_(Ixy2(0), KeyB) Then
    Swap_ = Ixy2
Else
    PushI Swap_, Ixy2(1)
    PushI Swap_, Ixy2(0)
End If
End Function
