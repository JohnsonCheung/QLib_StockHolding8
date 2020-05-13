Attribute VB_Name = "MxVbDtaDicIs"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Dic"
Const CMod$ = CLib & "MxVbDtaDicIs."

'- IsDikXxx
Function IsDikNm(A As Dictionary) As Boolean
Dim K: For Each K In A.Keys
    If Not IsNm(K) Then Exit Function
Next
IsDikNm = True
End Function
Function IsDikStr(D As Dictionary) As Boolean
IsDikStr = IsSyItr(D.Keys)
End Function

Function IsDicEmp(D As Dictionary) As Boolean
IsDicEmp = D.Count = 0
End Function

Function IsDiiLy(D As Dictionary) As Boolean
IsDiiLy = IsDiiSy(D)
End Function

Function IsDiiSy(D As Dictionary) As Boolean
Select Case True
Case Not IsSyItr(D.Items), Not IsItrStr(D.Keys)
Case Else: IsDiiSy = True
End Select
End Function

Function IsDiiAy(D As Dictionary) As Boolean
If Not IsItrAy(D.Items) Then Exit Function
IsDiiAy = True
End Function

Sub ChkDicabSamKey(D As Dictionary, B As Dictionary, Fun$)
If Not IsDicabSamKey(D, B) Then Thw Fun, "Give 2 dictionary are not EqKey"
End Sub

Function IsDicabSamKey(D As Dictionary, B As Dictionary) As Boolean
If D.Count <> B.Count Then Exit Function
Dim K: For Each K In D.Keys
    If Not B.Exists(K) Then Exit Function
Next
IsDicabSamKey = True
End Function

Sub ChkDiiIsStr(D As Dictionary, Fun$)
If Not IsDiiStr(D) Then Thw Fun, "Given Dic is not StrDic"
End Sub
Sub ChkDiiIsSy(D As Dictionary, Fun$)
If Not IsDiiSy(D) Then Thw Fun, "Given Dic is not SyDic"
End Sub

Sub ChkDiiIsLines(D As Dictionary, Fun$)
If Not IsDiiLines(D) Then Thw Fun, "Given Dic is not LinesDic"
End Sub

'-- IsDiX
Function IsDiiLines(D As Dictionary) As Boolean
Select Case True
Case Not IsItrLines(D.Items), Not IsItrStr(D.Keys)
Case Else: IsDiiLines = True
End Select
End Function

Function IsDiiStr(D As Dictionary) As Boolean
If Not IsItrStr(D.Keys) Then Exit Function
IsDiiStr = IsItrStr(D.Items)
End Function

Function IsDiiPrim(D As Dictionary) As Boolean
If Not IsPrimItr(D.Keys) Then Exit Function
IsDiiPrim = IsPrimItr(D.Items)
End Function
