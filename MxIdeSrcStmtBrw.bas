Attribute VB_Name = "MxIdeSrcStmtBrw"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcStmtBrw."
Const CNs$ = "AliMth"
Sub RfhBwStmtzM(M As CodeModule)
Dim S$(): S = Src(M)
Dim Ix: For Each Ix In RevAy(Mthixy(S))
    W1RfhI M, S, Ix
Next
End Sub
Private Sub W1RfhI(M As CodeModule, Src$(), Ix)
Dim L&: L = W1StmtLno(Src, Ix)
Dim Mthln$: Mthln = Contln(Src, Ix)
W1EnsMdln M, L, W1Stmt(Mthln)
End Sub
Private Function W1StmtLno&(Src$(), Ix)

End Function
Private Function W1Stmt$(Mthln$)

End Function
Private Sub W1EnsMdln(M As CodeModule, Lno&, Ln$)

End Sub
