Attribute VB_Name = "MxVbFsS32Atr"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Fs"
Const CMod$ = CLib & "MxVbFsS32Atr."

Type S32Atr
    Atr As String
    Val As String
End Type

Function S32AtrSi&(A() As S32Atr)
On Error Resume Next
S32AtrSi = UBound(A) + 1
End Function

Function S32AtrUB&(A() As S32Atr)
S32AtrUB = S32AtrSi(A) - 1
End Function

Sub PushS32Atr(O() As S32Atr, M As S32Atr)
Dim N&: N = S32AtrSi(O)
ReDim Preserve O(N)
O(N) = M
End Sub
