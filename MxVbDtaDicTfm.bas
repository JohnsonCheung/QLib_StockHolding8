Attribute VB_Name = "MxVbDtaDicTfm"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxVbDtaDicTfm."

Function DrszDic(A As Dictionary, Optional InlValTy As Boolean, Optional Tit$ = "Key Val") As Drs
DrszDic = DrszFF(Tit & " " & IIf(InlValTy, "ValTy", ""), DyzDic(A, InlValTy))
End Function

Function WszDic(Dic As Dictionary, Optional InlValTy As Boolean, Optional Tit$ = "Key Val") As Worksheet
Set WszDic = WszDrs(DrszDic(Dic, InlValTy))
End Function


Function DyzDic(A As Dictionary, Optional InlValTy As Boolean) As Variant()
Dim I, Dr
If A.Count = 0 Then Exit Function
Dim K(): K = A.Keys
If Si(K) = 0 Then Exit Function
For Each I In K
    If InlValTy Then
        Dr = Array(I, A(I), TypeName(A(I)))
    Else
        Dr = Array(I, A(I))
    End If
    Push DyzDic, Dr
Next
End Function
