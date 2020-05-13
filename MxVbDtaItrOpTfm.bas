Attribute VB_Name = "MxVbDtaItrOpTfm"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ns"
Const CMod$ = CLib & "MxVbDtaItrOpTfm."

Function IntozAy(Into, Ay)
If IsEqTy(Ay, Into) Then
    IntozAy = Ay
    Exit Function
End If
IntozAy = NwAy(Into)
Dim I
For Each I In Itr(Ay)
    Push IntozAy, I
Next
End Function

Function IntozItrNy(Into$, Itr, Ny$())
Dim O: O = Into: Erase O
Dim Obj
For Each Obj In Itr
    If HasEle(Ny, Objn(Obj)) Then
        PushObj O, Obj
    End If
Next
IntozItrNy = O
End Function

Function AvzItr(Itr) As Variant()
AvzItr = IntozItr(EmpAv, Itr)
End Function

Function SyzItrP(Itr, Prpp) As String()
Dim I: For Each I In Itr
    PushI SyzItrP, Opv(I, Prpp)
Next
End Function

Function SyzItv(Itr) As String()
Dim I: For Each I In Itr
    PushI SyzItv, I.Value
Next
End Function

Function SyzItr(Itr) As String()
SyzItr = IntozItr(EmpSy, Itr)
End Function

Function IntozItr(Into, Itr)
Dim O: O = Into: Erase O
Dim V: For Each V In Itr
    Push O, V
Next
IntozItr = O
End Function

Function IntozItrm(Into, Itr, Map$)
Dim O: O = Into: Erase Into
Dim X: For Each X In Itr
    Push O, Run(Map, X)
Next
IntozItrm = O
End Function

Function SyzItp(Itr, Prpp) As String()
SyzItp = IntozItp(EmpSy, Itr, Prpp)
End Function

Function IntAyzItp(Itr, Prpp) As Integer()
IntAyzItp = IntozItp(EmpIntAy, Itr, Prpp)
End Function

Function IntozItp(Into, Itr, Prpp)
IntozItp = NwAy(Into)
Dim Obj: For Each Obj In Itr
    Push IntozItp, Opv(Obj, Prpp)
Next
End Function

Function MaxItp(Itr, Prpp$)
Dim O, M
Dim Obj: For Each Obj In Itr
    M = Opv(Obj, Prpp)
    If M > O Then O = M
Next
MaxItp = O
End Function
Function AvzItp(Itr, P$) As Variant()
AvzItp = IntozItp(EmpAv, Itr, P)
End Function
