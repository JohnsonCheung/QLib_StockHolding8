Attribute VB_Name = "MxVbAyWdt"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CNs$ = "Wdty"
Const CMod$ = CLib & "MxVbAyWdt."

Private Function WdtyzDy_AllCol(Dy()) As Integer()
Dim J%: For J = 0 To NColzDy(Dy) - 1
    PushI WdtyzDy_AllCol, WdtzAy(StrColzDy(Dy, J))
Next
End Function

Private Function WdtyzDy_FstNCol(Dy(), FstNCol%) As Integer()
ReDim O(FstNCol - 1)
Dim J%: For J = 0 To FstNCol - 1
    PushI WdtyzDy_FstNCol, WdtzAy(StrColzDy(Dy, J))
Next
End Function

Function WdtyzDy(Dy(), Optional FstNCol%) As Integer()
If FstNCol <= 0 Then
    WdtyzDy = WdtyzDy_AllCol(Dy)
Else
    WdtyzDy = WdtyzDy_FstNCol(Dy, FstNCol)
End If
End Function
