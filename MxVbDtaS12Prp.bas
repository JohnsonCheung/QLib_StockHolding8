Attribute VB_Name = "MxVbDtaS12Prp"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbDtaS12Prp."

Function FstS2(S1, A() As S12) As StrOpt
'Ret : Fnd S1 in A return S2 @@
Dim J&: For J = 0 To S12UB(A)
    With A(J)
        If .S1 = S1 Then FstS2 = SomStr(.S2): Exit Function
    End With
Next
End Function


Function IsS12Lines(A As S12) As Boolean
Select Case True
Case IsLines(A.S1), IsLines(A.S2): IsS12Lines = True: Exit Function
End Select
End Function

Function IsS12yLines(A() As S12) As Boolean
Dim J&: For J = 0 To S12UB(A)
    If IsS12Lines(A(J)) Then IsS12yLines = True: Exit Function
Next
End Function
