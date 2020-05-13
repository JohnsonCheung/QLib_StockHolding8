Attribute VB_Name = "MxVbAyEle"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Ay"
Const CMod$ = CLib & "MxVbAyEle."
Function FstEle(Ay)
If Si(Ay) = 0 Then Exit Function
Asg Ay(0), _
    FstEle
End Function

Function LasEle(Ay)
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
Asg Ay(N - 1), _
    LasEle
End Function

Function MinEle(Ay)
If Si(Ay) = 0 Then Exit Function
Dim O: O = FstEle(Ay)
Dim I: For Each I In Itr(Ay)
    If I < O Then O = I
Next
MinEle = O
End Function

Function MaxEle(Ay)
If Si(Ay) = 0 Then Exit Function
Dim O: O = FstEle(Ay)
Dim I: For Each I In Itr(Ay)
    If I > O Then O = I
Next
MaxEle = O
End Function

Function LasSndEle(Ay)
Const CSub$ = CMod & "LasSndEle"
Dim N&: N = Si(Ay)
If N <= 1 Then
    Thw CSub, "Only 1 or no ele in Ay"
Else
    Asg Ay(N - 2), LasSndEle
End If
End Function
