Attribute VB_Name = "MxVbAyFstEle"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbAyFstEle."

Function RmvFstEle(Ay)
Const CSub$ = CMod & "RmvFstEle"
If Si(Ay) = 0 Then Thw CSub, "No FstEle"
Dim O: O = Ay: Erase O
Dim J&: For J = 1 To UB(Ay)
    PushI O, Ay(J)
Next
RmvFstEle = O
End Function

Function FstEleInAet(Ay, InAet As Dictionary)
Dim I
For Each I In Ay
    If InAet.Exists(I) Then FstEleInAet = I: Exit Function
Next
End Function

Function FstEleLik$(Ay, Lik$)
Dim X: For Each X In Itr(Ay)
    If X Like Lik Then FstEleLik = X: Exit Function
Next
End Function

Function FstElePredPX(Ay, PX$, P)
Dim X: For Each X In Itr(Ay)
    If Run(PX, P, X) Then
        Asg FstElePredPX, _
            X
        Exit Function
    End If
Next
End Function

Function FstElePredXABTrue(Ay, XAB$, A, B)
Dim X
For Each X In Itr(Ay)
    If Run(XAB, X, A, B) Then
        Asg FstElePredXABTrue, _
            X
        Exit Function
    End If
Next
End Function

Function FstElePredXP(A, XP$, P)
If Si(A) = 0 Then Exit Function
Dim X
For Each X In Itr(A)
    If Run(XP, X, P) Then
        Asg FstElePredXP, _
            X
        Exit Function
    End If
Next
End Function

Function FstElewRmvT1$(Sy$(), T1)
FstElewRmvT1 = RmvT1(FstElewT1(Sy, T1))
End Function

Function FstElewT1$(Ay, T1)
Dim I
For Each I In Itr(Ay)
    If FstTerm(I) = T1 Then FstElewT1 = I: Exit Function
Next
End Function

Function FstPfx$(Pfxy$(), S)
Dim P: For Each P In Pfxy
    If HasPfx(S, P) Then FstPfx = P: Exit Function
Next
End Function

Function FstElezRmvT1$(Sy$(), T1)
FstElezRmvT1 = RmvT1(FstElezT1(Sy, T1))
End Function

Function FstElezT1$(Sy$(), T1)
Dim S: For Each S In Itr(Sy)
    If HasT1(S, T1) Then FstElezT1 = S: Exit Function
Next
End Function

Function FstElezT2$(Sy$(), T2)
Dim S: For Each S In Itr(Sy)
    If HasT2(S, T2) Then FstElezT2 = S: Exit Function
Next
End Function

Function FstElezTT$(Sy$(), T1, T2)
Dim I, S$
For Each I In Itr(Sy)
    S = I
    If HasTT(S, T1, T2) Then FstElezTT = S: Exit Function
Next
End Function

Function FstNEle(Ay, N)
FstNEle = AwEix(Ay, N - 1)
End Function

Function ShfFstEle(OAy)
ShfFstEle = FstEle(OAy)
OAy = AeFstNEle(OAy)
End Function
