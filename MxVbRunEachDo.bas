Attribute VB_Name = "MxVbRunEachDo"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Ay"
Const CMod$ = CLib & "MxVbRunEachDo."

Sub EachDo(Ay, FunNm$)
Dim X: For Each X In Itr(Ay)
    Run FunNm, X
Next
End Sub

Sub EachDoABX(Ay, ABX$, A, B)
Dim X: For Each X In Itr(Ay)
    Run ABX, A, B, X
Next
End Sub

Sub EachDoAXB(Ay, AXB$, A, B)
Dim X: For Each X In Itr(Ay)
    Run AXB, A, X, B
Next
End Sub

Sub EachDoPPXP(A, PPXP$, P1, P2, P3)
Dim X: For Each X In Itr(A)
    Run PPXP, P1, P2, X, P3
Next
End Sub

Sub EachDoPX(A, PX$, P)
Dim X: For Each X In Itr(A)
    Run PX, P, X
Next
End Sub

Sub EachDoXP(A, XP$, P)
Dim X: For Each X In Itr(A)
    Run XP, X, P
Next
End Sub
