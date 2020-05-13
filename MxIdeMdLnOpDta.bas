Attribute VB_Name = "MxIdeMdLnOpDta"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdLnOpDta."
Public Const MdLnOpFF = "Mdn OpLno LinOp OldL NewL" ' ::Mdn *MdyLn:
Type MdLnOpRec
    Mdn As String
    OpLno As Long
    LinOp As String
    Oldl As String
    Newl As String
End Type

Function MdLnOpRec(Dr) As MdLnOpRec
With MdLnOpRec
    .Mdn = Dr(0)
    .OpLno = Dr(1)
    .LinOp = Dr(2)
    .Oldl = Dr(3)
    .Newl = Dr(4)
End With
End Function

Sub ChkIsMdMdyLnDrs(Fun$, D As Drs)
ChkDrsFF Fun, "MdMdyLn", D, MdLnOpFF
End Sub

Sub ClrMd(M As CodeModule)
If M.CountOfLines > 0 Then
    M.DeleteLines 1, M.CountOfLines
End If
End Sub

Sub DltLines(M As CodeModule, Lno&, OldLines$)
If OldLines = "" Then Exit Sub
If Lno = 0 Then Exit Sub
Dim Cnt&: Cnt = LnCnt(OldLines)
If M.Lines(Lno, Cnt) <> OldLines Then Thw CSub, "OldL <> ActL", "OldL ActL", OldLines, M.Lines(Lno, Cnt)
Debug.Print FmtQQ("DltLines: Lno(?) Cnt(?)", Lno, Cnt)
D Box(SplitCrLf(OldLines))
D ""
M.DeleteLines Lno, Cnt
End Sub

Sub InsLines(M As CodeModule, Lno, Lines$)
M.InsertLines Lno, Lines
Const CSub$ = CMod & "DltLinzD"
End Sub

Sub DltLinzD(M As CodeModule, L_OldL As Drs)
If JnSpc(L_OldL.Fny) <> "L OldL" Then Stop: Exit Sub
Stop
Dim B As Drs: B = SrtDrs(L_OldL, "-L")
Dim Dr
Stop
For Each Dr In Itr(B.Dy)
    Dim L&: L = Dr(0)
    Dim Oldl$: Oldl = Dr(2)
    Dim Newl$: Newl = Dr(1)
    If M.Lines(L, 1) <> Oldl Then Thw CSub, "Md-Ln <> OldL", "Mdn Lno Md-Ln OldL NewL", Mdn(M), L, M.Lines(L, 1), Oldl, Newl
    M.ReplaceLine L, Newl
Next
End Sub

Sub InsLinzD(M As CodeModule, L_NewL As Drs)
If JnSpc(L_NewL.Fny) <> "L NewL" Then Stop: Exit Sub
Stop
Dim B As Drs: B = SrtDrs(L_NewL, "-L")
Dim Dr
Stop
For Each Dr In Itr(B.Dy)
    Dim L&: L = Dr(0)
    Dim Oldl$: Oldl = Dr(2)
    Dim Newl$: Newl = Dr(1)
    If M.Lines(L, 1) <> Oldl Then Thw CSub, "Md-Ln <> OldL", "Mdn Lno Md-Ln OldL NewL", Mdn(M), L, M.Lines(L, 1), Oldl, Newl
    M.ReplaceLine L, Newl
Next
End Sub

Function LasDclLno&(M As CodeModule)
LasDclLno = M.CountOfDeclarationLines
End Function

Sub AppDcl(M As CodeModule, Lines$)
M.InsertLines LasDclLno(M) + 1, Lines
End Sub

Sub InsLin(M As CodeModule, L_NewL As Drs)
Dim B As Drs: B = L_NewL
If JnSpc(B.Fny) <> "L NewL" Then Stop: Exit Sub
Dim Dr
For Each Dr In Itr(B.Dy)
    Dim L&: L = Dr(0)
    Dim Newl$: Newl = Dr(1)
    M.InsertLines L, Newl
Next
End Sub

Sub MdyLnzUpd(WiMdMdyLn As Drs, Upd As eUpdRpt)
If IsRpt(Upd) Then BrwDrs WiMdMdyLn, FnPfx:="MdMdyLn_"
If IsUpd(Upd) Then MdyLn WiMdMdyLn
End Sub

Sub MdyLn(WiMdMdyLn As Drs)
Dim G As Drs: G = GDrs(WiMdMdyLn, "Mdn", "OpLno LinOp NewL OldL")
Dim Fny$(): Fny = SyzSS(MdyLnFF)
Dim Dr: For Each Dr In Itr(G.Dy)
    Dim IMdn$: IMdn = Dr(0)
    Dim IMd As CodeModule: Set IMd = Md(IMdn)
    Dim IDy(): IDy = Dr(1)
    Dim IDoMdyLn As Drs: IDoMdyLn = Drs(Fny, IDy)
    MdyLnzM IMd, IDoMdyLn
Next
End Sub

Sub ChkDrsFF(Fun$, FFoXXX$, D As Drs, EptFF$)
If FF(D) <> EptFF Then Thw Fun, "Given @GivenFF <> @EptFF", "@FFoXX GivenFF @EptFF", FFoXXX, FF(D), EptFF
End Sub

Sub ChktDoMdyLn(Fun$, D As Drs)
ChkDrsFF Fun, "MdyLn", D, MdyLnFF
End Sub

Sub RplLNewOzM(M As CodeModule, L&, Oldl$, Newl$)
Dim A_NowL$: A_NowL = M.Lines(L, 1)
If A_NowL <> Oldl Then ThwRplLNewOMisMch CSub, Mdn(M), L, Oldl, A_NowL, Newl
M.ReplaceLine L, Newl '<====
Debug.Print Tab; "Rpl"; L; "Old --->"; Oldl
Debug.Print Tab; Space(Len("Rpl" & L)); "New --->"; Newl
End Sub

Sub DltLinzM(M As CodeModule, L&, Oldl$)
Dim A_NowL$: A_NowL = M.Lines(L, 1)
If A_NowL <> Oldl Then
    Const A_Msg1$ = "The Md-Line going to DELETE is not expected"
    Const A_NN1$ = "Mdn Lno [Should be this line] [But Md has this line]"
    Thw CSub, A_Msg1, A_NN1, Mdn(M), L, Oldl, A_NowL
End If
M.DeleteLines L, 1 '<====
Debug.Print Tab; "Dlt"; L; "--->"; Oldl
End Sub

Sub InsLinzM(M As CodeModule, L&, Newl$)
M.InsertLines L, Newl  '<===
Debug.Print Tab; "Ins"; L; "--->"; Newl
End Sub

Sub MdyLnzM(M As CodeModule, DoMdyLn As Drs)
'If MsgBox("Going to modify the above module.", vbOKCancel, Mdn(M)) = VbMsgBoxResult.vbCancel Then Stop
If Mdn(M) = "MxMdMdy" Then
    Debug.Print "MdyLnzM: Md-> MxMdMdy <===================== Skipped"
    Exit Sub
End If
Debug.Print "MdyLnzM: Md->"; Mdn(M)
ChktDoMdyLn CSub, DoMdyLn
Dim D As Drs: D = SrtDrs(DoMdyLn, "-OpLno LinOp")
'Insp CSub, "Check given @DoMdyLn :DoMdyLno match given given @Md #Src", "MdyLnFF @DoMdyLn @Mdn Src", MdyLnFF, FmtDrs(D), Mdn(M), AddIxPfx(Src(M), eWsAtBeg1): Stop
Dim IL&, ILinOp$, INewL$, IOldL$, INowL$ '<-- Each Dr in WiL_LinOp_NewL_OldL
Dim Dr: For Each Dr In Itr(D.Dy)
    IL = Dr(0)
    ILinOp = Dr(1)
    INewL = Dr(2)
    IOldL = Dr(3) '
    Select Case ILinOp
    Case "Rpl": RplLNewOzM M, IL, IOldL, INewL
    Case "Dlt": DltLinzM M, IL, IOldL
    Case "Ins": InsLinzM M, IL, INewL
    Case Else:  Thw CSub, "Given Drs->LinOp is invalid.  Valid LinOp=[Rpl Dlt Ins]", "Drs_L_LinOp_NewL_OldL", FmtDrsR(D)
    End Select
Next
End Sub

Private Sub ThwRplLNewOMIsMch__Tst()
ThwRplLNewOMisMch "MdyLnzM", "Md-A", 123, "OldL", "NowL", "NewL"
End Sub

Sub ThwRplLNewOMisMch(Fun$, Mdn$, L&, Oldl$, NowL$, Newl$)
Const CSub$ = CMod & "RplLNewO"
Const WMsg$ = "The Md-Line going to REPLACE is not expected"
Const WNN$ = "Mdn Lno [Should be this line] [But Md has this line] [The new line]"
Thw Fun, WMsg, WNN, Mdn, L, Oldl, NowL, Newl
End Sub


Private Function MsgzDoMdy__N%(WiLinOp As Drs, LinOp$)
MsgzDoMdy__N = NReczDrs(DwEq(WiLinOp, "LinOp", LinOp))
End Function

Function MsgzDoMdyLn(L_LinOp_OldL_NewL As Drs, Msg$) As String()
Dim D As Drs: D = L_LinOp_OldL_NewL
Dim NRpl%: NRpl = MsgzDoMdy__N(D, "Rpl")
Dim NIns%: NIns = MsgzDoMdy__N(D, "Ins")
Dim NDlt%: NDlt = MsgzDoMdy__N(D, "Dlt")
Dim CntMsg$:  CntMsg = FmtQQ("NRpl(?) NDlt(?) NIns(?)", NRpl, NDlt, NIns)
PushI MsgzDoMdyLn, Msg
PushI MsgzDoMdyLn, ""
PushIAy MsgzDoMdyLn, FmtNav(Av("Cnt DoMdyLn", CntMsg, FmtDrsR(L_LinOp_OldL_NewL)))
End Function
