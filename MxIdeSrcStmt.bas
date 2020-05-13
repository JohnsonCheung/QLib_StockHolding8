Attribute VB_Name = "MxIdeSrcStmt"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcStmt."

Sub Stmt__Tst()
BrwAy Stmt(SrcM), "Stmt__Tst"
End Sub

Function Stmty(Srcy()) As Variant()
Dim S: For Each S In Srcy
    PushI Stmty, Stmt(CvSy(S))
Next
End Function

Function RmvStmtRmk(Stmt$()) As String()
Dim L: For Each L In Itr(Stmt)
    PushNB RmvStmtRmk, BrkVrmk(L).S1
Next
End Function

Function Stmt(Src$()) As String()
Dim L: For Each L In Itr(Contly(Src))
    PushIAy Stmt, StmtzL(L)
Next
End Function

Private Sub StmtzL__Tst()
Dim Contln$
'GoSub T1
GoSub T2
Exit Sub
T1:
    Contln = "Dim A$: B"
    Ept = Sy("Dim A$", "B")
    Pass "T1 Brk"
    GoTo Tst
T2:
    Contln = "Label: AAA"
    Ept = Sy("Label:", "AAA")
    Pass "T1 Label"
    GoTo Tst
Tst:
    Act = StmtzL(Contln)
    C
    Return
End Sub

Function StmtzL(Contln) As String()
If Contln = "" Then Exit Function
Dim Rmk$, Ln$
    Dim S As S12: S = BrkVrmk(Contln)
    If S.S2 <> "" Then Rmk = "' " & S.S2
    Ln = S.S1
If Ln = "" Then If Rmk <> "" Then PushI StmtzL, Rmk: Exit Function

Dim O$()
    O = AmTrim(SplitPosy(S.S1, StmtBrkColonPosy(S.S1)))
    If Rmk <> "" Then
        Dim U&: U = UB(O)
        O(U) = O(U) & Rmk
    End If
StmtzL = O
End Function

Function StmtBrkColonPosy(NoRmkContln) As Integer()
Dim Py%(): Py = SubStrPosy(NoRmkContln, ":")
Dim IsFstColon As Boolean: IsFstColon = True
Dim P: For Each P In Itr(Py)
    If IsStmtBrkColon(NoRmkContln, P, IsFstColon) Then PushI StmtBrkColonPosy, P
Next
End Function

Function IsStmtBrkColon(NoRmkContln, Pos, OIsFstColon As Boolean) As Boolean
If OIsFstColon Then
    OIsFstColon = False
    If IsNm(Left(NoRmkContln, Pos - 1)) Then Exit Function 'It is a label-Colon, not Stmtbrk-Colon
End If
Dim C%: C = DblQCnt(Left(NoRmkContln, Pos - 1))
IsStmtBrkColon = IsEven(C)
End Function
