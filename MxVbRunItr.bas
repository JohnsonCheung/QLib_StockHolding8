Attribute VB_Name = "MxVbRunItr"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Itr"
Const CMod$ = CLib & "MxVbRunItr."

Sub ItrDo(Itr, DoFun$)
Dim I: For Each I In Itr
    Run DoFun, I
Next
End Sub

Sub ItrDoPX(Itr, PX$, P)
Dim X: For Each X In Itr
    Run PX, P, X
Next
End Sub

Sub ItrDoXP(Itr, XP$, P)
Dim X: For Each X In Itr
    Run XP, X, P
Next
End Sub
