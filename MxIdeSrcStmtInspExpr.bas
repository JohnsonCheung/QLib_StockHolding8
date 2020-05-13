Attribute VB_Name = "MxIdeSrcStmtInspExpr"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcStmtInspExpr."
'==InspStmt
Private Sub InspStmtzL__Tst()
Dim O() As S12, J%
Dim C As VBComponent: For Each C In CPj.VBComponents
    PushS12 O, S12(C.Name, JnCrLf(InspStmty(Mthlny(Src(C.CodeModule)), C.Name)))
    J = J + 1
    If J > 10 Then Exit For
Next
BrwS12y O
End Sub
Private Function InspStmty(Mthlny$(), Mdn$) As String()
Dim Mthln: For Each Mthln In Itr(Mthlny)
    PushNB InspStmty, InspStmtzL(Mthln, Mdn)
Next
End Function
Function InspStmtzL$(Mthln, Mdn$)
With MsigzL(Mthln)
    Dim IsRet As Boolean: IsRet = IsRetVal(.ShtTy)
    If ArgSi(.Arg) = 0 And Not IsRet Then Exit Function
    Dim NN$:  NN = W1aNN(.Arg, IsRet)
    Dim Epr$: Epr = W1aEpr(.Arg, IsRet, .Ret)
    Dim Fun$: Fun = Mdn & "." & .Mthn
    InspStmtzL = InspStmt(Fun, NN, Epr)
End With
End Function
Function InspStmt$(Fun$, Varnn$, EprLis$, Optional Msg = "Inspect")
Const C$ = "Insp ""?"", ""?"", ""?"", ?"
InspStmt = FmtQQ(C, Fun, Msg, Varnn, EprLis)
End Function
Private Function W1aNN$(A() As Arg, IsRet As Boolean)
Dim N$: If IsRet Then N = "Ret "
W1aNN = N & JnSpc(ArgNy(A))
End Function
Private Function W1aEpr$(Arg() As Arg, IsRet As Boolean, RetVty As Vty)
Dim O$(): If IsRet Then PushI O, W1EprzRet(IsRet, RetVty) '#Insp-Epr-0
Dim J%: For J = 0 To ArgUB(Arg)
    PushI O, W1Epr(Arg(J))
Next
W1aEpr = JnCommaSpc(O)
End Function
Private Function W1Epr$(A As Arg)
If W1IsStrbVty(A.Ty) Then W1Epr = A.Argn: Exit Function
W1Epr = W1ToStrEpr(A.Argn, A.Ty)
End Function
Private Function W1IsStrbVty(A As Vty) As Boolean '#Is-Stringable-Vty# is the Vty can be expressed in an expression
If A.TyChr <> "" Then Exit Function
End Function
Private Function W1EprzRet$(IsRetVal As Boolean, Ret As Vty)
If IsRetVal Then W1EprzRet = W1ToStrEpr("Ret", Ret)
End Function
Private Function W1ToStrEpr$(Nm$, T As Vty)
Dim O$
If T.TyChr <> "" Then W1ToStrEpr = O: Exit Function
Select Case T.Tyn
Case "Drs":        O = Bktv("FmtDrs", Nm)
Case "S12":        O = Bktv("FmtS12" & IIf(T.IsAy, "y", ""), Nm)
Case "CodeModule": O = Bktv("Mdn", Nm)
Case "Dictionary": O = Bktv("FmtDic", Nm)
Case Else: O = """NoFmtr(" & T.Tyn & ")"""
End Select
W1ToStrEpr = O
End Function
Function Bktv$(BefBktn$, InsidBktn$): Bktv = BefBktn & "(" & InsidBktn & ")": End Function
