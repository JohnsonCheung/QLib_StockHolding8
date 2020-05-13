Attribute VB_Name = "MxIdeSrcGenLib"
Option Compare Text
Option Explicit
Const CNs$ = "Gen"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcGenLib."

Sub GenFxLib()
GenFxLibzP CPj
End Sub

Sub GenFbLib()
GenFbLibzP CPj
End Sub

Sub GenFxLibzP(P As VBProject)
Dim I: For Each I In Itr(LibItmAy(P))
    ExpLibItm P, I
    GenFba SrcPthzLibItm(I)
Next
End Sub

Function SrcPthzLibItm$(LibItm)

End Function

Sub ExpLibItm(P As VBProject, LibItm)

End Sub

Function LibItmAy(P As VBProject) As String()

End Function

Sub GenFbLibzP(P As VBProject)

End Sub
