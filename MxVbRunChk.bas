Attribute VB_Name = "MxVbRunChk"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Thw"
Const CMod$ = CLib & "MxVbRunChk."

Sub ChkNoEr(Er$(), Fun$)
If Si(Er) = 0 Then Exit Sub
Thw Fun, JnCrLf(Er)
End Sub

Sub ChkIsStr(A, Fun$)
If IsStr(A) Then Exit Sub
Thw Fun, "Given parameter should be str, but now TypeName=" & TypeName(A)
End Sub
Sub ChkTrue(ShouldTrue As Boolean, Msg$)
If Not ShouldTrue Then Raise Msg
End Sub

Sub ChkEq(A, B, Optional N12$ = "A B")
Const CSub$ = CMod & "ChkEq"
ChkSamTy A, B, N12
Select Case True
Case IsLines(A), IsLines(B)
                  If A <> B Then CprLines CStr(A), CStr(B), Hdr:=FmtQQ("Lines of name12[?] not eq.", N12): Stop: Exit Sub
Case IsStr(A):    If A <> B Then ChkIsEqStr CStr(A), CStr(B), Hdr:=FmtQQ("String of name12[?] not eq.", N12): Stop: Exit Sub
Case IsDic(A):    If Not IsEqDic(CvDic(A), CvDic(B)) Then BrwCprDic CvDic(A), CvDic(B): Stop: Exit Sub
Case IsArray(A):  ChkEqAy A, B, N12, CSub
Case IsObject(A): If ObjPtr(A) <> ObjPtr(B) Then Thw CSub, "Two values are diff type", "N12 Tyn1 Tyn2", N12, TypeName(A), TypeName(B)
Case Else:
    If A <> B Then
        Thw CSub, "A B NE", "N12 A B", N12, A, B
        Exit Sub
    End If
End Select
End Sub
Sub ChkEqSi(AyA, AyB, Optional Fun$ = "ChkEqAy")
If Si(AyA) <> Si(AyB) Then Raise Fun & ": Two array are dif size"
End Sub

Sub ChkEqAy(AyA, AyB, Optional Ayn2$ = "Ay1 Ay2", Optional Fun$ = "ChkEqAy")
Dim N As S12: N = Brk1Spc(Ayn2)
ChkIsAy AyA, N.S1, Fun
ChkSamSi AyA, AyB, Ayn2, Fun
ChkSamTy AyA, AyB, Ayn2, Fun

Dim J&, A
For Each A In Itr(AyA)
    If Not IsEq(A, AyB(J)) Then
        Dim NN$: NN = "AyN2 AyTypeName Dif-Ix Dif-V1Ty Dif-V2Ty Dif-V1 Dif-V2 Ay1 Ay2"
        Thw Fun, "There is ele in 2 Ay are diff", NN, Ayn2, TypeName(AyA), J, TypeName(A), TypeName(AyB(J)), A, AyB(J), AyA, AyB
        Exit Sub
    End If
    J = J + 1
Next
End Sub

Sub ChkSamTy(A, B, Optional Nm2$ = "A B", Optional Fun$ = "ChkSamTy")
If TypeName(A) = TypeName(B) Then Exit Sub
Dim N$
With Brk1Spc(Nm2)
    N = FmtQQ("?/?", .S1, .S2)
    Thw Fun, "TypeName of 2 var are Diff:", "Nm Ty1 Ty2", N, TypeName(A), TypeName(B)
End With
End Sub

Sub ChkSamSi(Ay1, Ay2, Optional Ayn2$ = "Ay1 Ay2", Optional Fun$ = "ChkSamSi")
Dim A As S12: A = BrkSpc(Ayn2)
If Si(Ay1) <> Si(Ay2) Then Thw Fun, FmtQQ("Si-of-[?]-[?] <> Si-of-[?]-[?]", A.S1, Si(Ay1), A.S2, Si(Ay2))
End Sub

Sub ChkHasFF(A As Drs, FF$, Fun$)
If JnSpc(A.Fny) <> FF Then Thw Fun, "Drs-FF <> FF", "Drs-FF FF", JnSpc(A.Fny), FF
End Sub

Sub ChkSrt(Ay, Fun$)
If IsSrtd(Ay) Then Thw Fun, "Array should be sorted", "Ay-Ty Ay", TypeName(Ay), Ay
End Sub

Sub ChkSomthing(A, VarNm$, Fun$)
If Not IsNothing(A) Then Exit Sub
Thw Fun, FmtQQ("Given[?] is nothing", VarNm)
End Sub

Sub ChkIsPrimy(A, Optional Fun$ = CMod & "ChkIsPrimy")
If IsPrimy(A) Then Exit Sub
Thw Fun, "Given parameter should be prim-array", "Tyn", TypeName(A)
End Sub

Function IsPrimy(Ay) As Boolean
If Not IsArray(Ay) Then Exit Function
IsPrimy = IsPrimTy(RmvLas2Chr(TypeName(Ay)))
End Function

Sub ChkIsAy(A, Optional AyNm$ = "Ay", Optional Fun$ = "ChkIsAy")
If IsEmpty(A) Then Exit Sub
If IsArray(A) Then Exit Sub
Thw Fun, "Given parameter should be array", "AyNm Tyn", AyNm, TypeName(A)
End Sub

Sub ChkDifObj(A, B, Fun$, Optional Msg$ = "Two given object cannot be same")
If IsEqObj(A, B) Then Thw Fun, Msg
End Sub

Sub ChkGoodIxAy(IxAy, Fun$)
Dim O$()
    Dim I, J&: For Each I In Itr(IxAy)
        If I < 0 Then
            PushI O, J & ": " & I
            J = J + 1
        End If
    Next
If Si(O) > 0 Then
    Thw Fun, "In [IxAy], there are [negative-element (Ix Ele)]", "IxAy Neg-Ele", IxAy, O
End If
End Sub
