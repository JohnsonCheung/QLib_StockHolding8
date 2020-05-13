Attribute VB_Name = "MxVbVarIsPrimSy1"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbVarIsPrimSy1."


Function IsDteSy(Sy$()) As Boolean
Dim S: For Each S In Sy
    If Not IsDteStr(S) Then Exit Function
Next
IsDteSy = True
End Function

Function IsDblSy(Sy$()) As Boolean
Dim S: For Each S In Sy
    If Not IsDblStr(S) Then Exit Function
Next
IsDblSy = True
End Function
