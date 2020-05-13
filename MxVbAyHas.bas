Attribute VB_Name = "MxVbAyHas"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay.Has"
Const CMod$ = CLib & "MxVbAyHas."
Function HasObj(ObjAy, Obj) As Boolean
Dim OPtr&: OPtr = ObjPtr(Obj)
Dim I: For Each I In ObjAy
    If ObjPtr(I) = OPtr Then HasObj = True: Exit Function
Next
End Function

Function HasDup(Ay) As Boolean
Dim S As New Dictionary
Dim I: For Each I In Ay
    If S.Exists(I) Then HasDup = True: Exit Function
    S.Add I, Empty
Next
End Function

Function HasStrEle(Ay, StrEle$, Optional C As VbCompareMethod = vbBinaryCompare) As Boolean
Dim I: For Each I In Itr(Ay)
    If IsEqStr(I, StrEle, C) Then HasStrEle = True: Exit Function
Next
End Function

Function HasEleRe(Ay, Re As RegExp) As Boolean
Dim Ele: For Each Ele In Itr(Ay)
    If Re.Test(Ele) Then HasEleRe = True: Exit Function
Next
End Function
Function SomEle(Ay) As Boolean
SomEle = Si(Ay) <> 0
End Function

Function IsEmpAy(Ay) As Boolean
IsEmpAy = Si(Ay) = 0
End Function

Function NoEle(Ay, Ele) As Boolean
NoEle = Not HasEle(Ay, Ele)
End Function

Function HasSubAy(Ay, SubAy) As Boolean
Dim S: For Each S In Itr(SubAy)
    If NoEle(Ay, S) Then Exit Function
Next
HasSubAy = True
End Function

Function HasEle(Ay, Ele) As Boolean
Dim I
For Each I In Itr(Ay)
    If I = Ele Then HasEle = True: Exit Function
Next
End Function

Function HasEleAy(Ay, EleAy) As Boolean
Dim I
For Each I In Itr(EleAy)
    If Not HasEle(Ay, I) Then Exit Function
Next
HasEleAy = True
End Function

Function HasElezInSomAyzOfAp(ParamArray AyAp()) As Boolean
Dim AvAp(): AvAp = AyAp
Dim Ay
For Each Ay In Itr(AvAp)
    If Si(Ay) > 0 Then HasElezInSomAyzOfAp = True: Exit Function
Next
End Function

Function IsAySub(SubAy, SuperAy) As Boolean
Dim I: For Each I In Itr(SubAy)
    If Not HasEle(SuperAy, I) Then Exit Function
Next
IsAySub = True
End Function

Function IsAySuper(SuperAy, SubAy) As Boolean
IsAySuper = IsAySub(SubAy, SuperAy)
End Function

Function ChktSuperSubAy(SuperAy, SubAy) As String()
Const CSub$ = CMod & "ChktSuperSubAy"
If IsAySuper(SuperAy, SubAy) Then Exit Function
Thw CSub, "Super-Sub-Ay error", "ErEle-in-SubAy SubAy SuperAy", MinusAy(SubAy, SuperAy), SubAy, SuperAy
End Function

Function HasEleAyInSeq(A, B) As Boolean
Dim BItm, Ix&
If Si(B) = 0 Then Stop
For Each BItm In B
    Ix = IxzAy(A, BItm, Ix)
    If Ix = -1 Then Exit Function
    Ix = Ix + 1
Next
HasEleAyInSeq = True
End Function

Function HasDupEle(A) As Boolean
If Si(A) = 0 Then Exit Function
Dim Pool: Pool = A: Erase Pool
Dim I
For Each I In A
    If HasEle(Pool, I) Then HasDupEle = True: Exit Function
    Push Pool, I
Next
End Function

Function HasNegEle(A) As Boolean
Dim V
If Si(A) = 0 Then Exit Function
For Each V In A
    If V = -1 Then HasNegEle = True: Exit Function
Next
End Function

Function HasElePredPXTrue(A, PX$, P) As Boolean
Dim X
For Each X In Itr(A)
    If Run(PX, P, X) Then HasElePredPXTrue = True: Exit Function
Next
End Function

Function HasElePredXPTrue(A, XP$, P) As Boolean
If Si(A) = 0 Then Exit Function
Dim X
For Each X In Itr(A)
    If Run(XP, X, P) Then
        HasElePredXPTrue = True
        Exit Function
    End If
Next
End Function

Private Sub HasEleAyInSeq__Tst()
Dim A, B
A = Array(1, 2, 3, 4, 5, 6, 7, 8)
B = Array(2, 4, 6)
Debug.Assert HasEleAyInSeq(A, B) = True
End Sub

Sub ChkEle(Ay, Ele, Fun$)
If NoEle(Ay, Ele) Then Thw Fun, "No @Ele in @Ay", "Ele Ay", Ele, Ay
End Sub
