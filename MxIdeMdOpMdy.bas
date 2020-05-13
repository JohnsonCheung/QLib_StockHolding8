Attribute VB_Name = "MxIdeMdOpMdy"
Option Explicit
Option Compare Text

Sub MdyMd()

End Sub
Sub DltQvv()
Dim C As VBComponent: For Each C In CPj.VBComponents
    DltQvvzM C.CodeModule
Next
End Sub
Private Sub DltQvvzM(M As CodeModule)
Dim J%: For J = M.CountOfLines To 1 Step -1
    If HasPfx(M.Lines(J, 1), "'^^") Then
        M.DeleteLines J, 1
        Debug.Print "Deleted: " & M.Parent.Name, J
    End If
Next
End Sub
Private Function HasPfx(S, Pfx$) As Boolean: HasPfx = Left(S, Len(Pfx)) = Pfx: End Function
Sub LisQvv()
Dim C As VBComponent: For Each C In CPj.VBComponents
    LisQvvzM C.CodeModule
Next
End Sub
Private Sub LisQvvzM(M As CodeModule)
Dim J%: For J = 1 To M.CountOfLines
    If HasPfx(M.Lines(J, 1), "'^^") Then
        Debug.Print M.Parent.Name, J
    End If
Next
End Sub
