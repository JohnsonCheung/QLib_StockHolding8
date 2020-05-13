Attribute VB_Name = "MxVbAyIs"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxVbAyIs."
Function IsSuperAy(SuperAy, SubAy) As Boolean
Dim SubI: For Each SubI In Itr(SubAy)
    If NoEle(SuperAy, SubI) Then Exit Function
Next
IsSuperAy = True
End Function
