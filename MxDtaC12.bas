Attribute VB_Name = "MxDtaC12"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Dta"
Const CMod$ = CLib & "MxDtaC12."
Type C12
    C1 As Integer 'Start from 1
    C2 As Integer
End Type
Function IsEmpC12(A As C12) As Boolean
With A
Select Case True
Case .C1 <= 0, .C2 <= 0, .C1 > .C2: IsEmpC12 = True: Exit Function
End Select
End With
End Function
Function C12(C1, C2) As C12
If C1 <= 0 Then PmEr CSub, "C1 Must >=1", "C1", C1
If C2 <= 0 Then PmEr CSub, "C2 Must >=1", "C2", C2
With C12
    .C1 = C1
    .C2 = C2
End With
End Function
Function C12Str$(A As C12)
With A
C12Str = "C12 " & .C1 & " " & .C2
End With
End Function

Function IsEqC12(A As C12, B As C12) As Boolean
With A
    If .C1 <> B.C1 Then Exit Function
    If .C2 <> B.C2 Then Exit Function
End With
IsEqC12 = True
End Function

Sub PushC12(O() As C12, M As C12)
Dim N&: N = C12Si(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Function C12UB&(A() As C12)
C12UB = C12Si(A) - 1
End Function
Function C12Si&(A() As C12)
On Error Resume Next
C12Si = UBound(A) + 1
End Function

Sub ChkC12Within(A_Must As C12, In_B As C12)
Dim A As C12: A = A_Must
Dim B As C12: B = In_B
If B.C1 < A.C1 Then Stop
If B.C2 < A.C2 Then Stop
End Sub
