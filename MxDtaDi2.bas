Attribute VB_Name = "MxDtaDi2"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Dic"
Const CMod$ = CLib & "MxDtaDi2."
Type Di2
    A As Dictionary
    B As Dictionary
End Type
Function Di2(A As Dictionary, B As Dictionary) As Di2
ChkSomthing A, "DicA", CSub
ChkSomthing B, "DicB", CSub
With Di2
    Set .A = A
    Set .B = B
End With
End Function
Function Di2zInKy(D As Dictionary, InKy) As Di2
Dim K, A As New Dictionary, B As New Dictionary
For Each K In D.Keys
    If HasEle(InKy, K) Then
        A.Add K, D(K)
    Else
        B.Add K, D(K)
    End If
Next
Di2zInKy = Di2(A, B)
End Function
