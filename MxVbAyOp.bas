Attribute VB_Name = "MxVbAyOp"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay.Op"
Const CMod$ = CLib & "MxVbAyOp."
Function MinusAyAp(Ay, ParamArray Ap())
Dim IAy, O
O = Ay
For Each IAy In Ap
    O = MinusAy(Ay, IAy)
    If Si(O) = 0 Then MinusAyAp = O: Exit Function
Next
MinusAyAp = O
End Function

Function AddItmAy(Itm, Ay)
Dim O: O = Ay: Erase O
PushI O, Itm
PushIAy O, Ay
AddItmAy = O
End Function
Function CvVy(Vy)
Const CSub$ = CMod & "CvVy"
Select Case True
Case IsStr(Vy): CvVy = SyzSS(CStr(Vy))
Case IsArray(Vy): CvVy = Vy
Case Else: Thw CSub, "VyzDicKK should either be string or array", "Vy-TypeName Vy", TypeName(Vy), Vy
End Select
End Function

Sub ChkNoDup(Ay, Optional N$ = "Ay", Optional Fun$ = "ChkNoDup")
' If there are 2 ele with same string (IgnCas), throw error
Dim Dup$()
    Dup = AwDup(Ay)
If Si(Dup) = 0 Then Exit Sub
Thw Fun, "There are dup in array", "AyNm Dup Ay", N, Dup, Ay
End Sub

Function MinElezGT0(Ay)
If Si(Ay) = 0 Then Exit Function
Dim O: O = Ay(0)
Dim V: For Each V In Ay
    If V > 0 Then
        If O = 0 Then
            O = V
        Else
            If V < O Then O = V
        End If
    End If
Next
MinElezGT0 = O
End Function

Function AySum#(NumAy)
Dim O#, V: For Each V In Itr(NumAy)
    O = O + V
Next
AySum = O
End Function

Function HasIntersect(A, B) As Boolean
Dim I: For Each I In Itr(A)
    If HasEle(B, I) Then HasIntersect = True: Exit Function
Next
End Function

Function RplAy(A, ByAy, Bix)
Dim O: O = A
Dim J&: For J = 0 To UB(ByAy)
    O(J + Bix) = ByAy(J)
Next
RplAy = O
End Function
Function IntersectAy(A, B)
IntersectAy = NwAy(A)
If Si(A) = 0 Then Exit Function
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If HasEle(B, V) Then PushI IntersectAy, V
Next
End Function

Function MinusSy(A$(), B$()) As String()
MinusSy = MinusAy(A, B)
End Function

Function MinusAy(A, B)
If Si(B) = 0 Then MinusAy = A: Exit Function
MinusAy = NwAy(A)
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If Not HasEle(B, V) Then
        PushI MinusAy, V
    End If
Next
End Function

Function AddSy(SyA$(), SyB$()) As String()
AddSy = AddAy(SyA, SyB)
End Function

Function FlatAp(ParamArray Itm_or_Ay_Ap())
Dim Av(): Av = Itm_or_Ay_Ap
If Si(Av) = 0 Then Thw CSub, "Given Itm_or_Ay_Ap should have at least one element"
FlatAp = EmpAyzV(Av(0))
Dim I: For Each I In Av
    If IsArray(I) Then
        PushAy FlatAp, I
    Else
        Push FlatAp, I
    End If
Next
End Function

Function EmpAyzV(PrimVal_or_Ay) ' ret an empty array from @PrimVal_or_Ay
If IsArray(PrimVal_or_Ay) Then
    EmpAyzV = NwAy(PrimVal_or_Ay)
Else
    Dim T As VbVarType: T = VarType(PrimVal_or_Ay)
    Select Case True
    Case T = vbString: EmpAyzV = EmpSy
    Case T = vbInteger: EmpAyzV = EmpIntAy
    Case T = vbLong: EmpAyzV = EmpLngAy
    Case T = vbEmpty: EmpAyzV = EmpAv
    Case T = vbBoolean: EmpAyzV = EmpBoolAy
    Case T = vbByte: EmpAyzV = EmpBytAy
    Case T = vbDouble: EmpAyzV = EmpDblAy
    Case T = vbDate: EmpAyzV = EmpDteAy
    Case T = vbCurrency: EmpAyzV = EmpCurAy
    
    Case Else: Thw CSub, "Given V is not Ay nor Prim", "TypeName-V", TypeName(PrimVal_or_Ay)
    End Select
End If
End Function
Function AddAy(AyA, AyB)
AddAy = AyA
PushAy AddAy, AyB
End Function

Function AddAv(A(), B()) As Variant()
AddAv = A
PushAy AddAv, B
End Function
