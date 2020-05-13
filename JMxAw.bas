Attribute VB_Name = "JMxAw"
Option Compare Text
Const CMod$ = CLib & "JMxAw."
#If False Then
Option Explicit

Function AwDup(Ay)
Dim O: O = Ay: Erase O
Dim I, J&: For Each I In Itr(Ay)
    If Not HasEle(O, I) Then
        If HasEleFm(Ay, I, J + 1) Then Push O, I: Exit For
    End If
    J = J + 1
Next
AwDup = O
End Function

#End If
