Attribute VB_Name = "JMxIsTy"
Option Compare Text
Const CMod$ = CLib & "JMxIsTy."
#If False Then
Option Explicit
Function IsSy(V) As Boolean
IsSy = VarType(V) = vbArray + vbString
End Function
#End If
