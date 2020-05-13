Attribute VB_Name = "MxIdeMthMsig"
Option Compare Text
Option Explicit
'**MthlnFun
Function BktzIsAy$(IsAy As Boolean): BktzIsAy = SzTrue(IsAy, "()"): End Function

'**Msig
Private Sub MsigzL__Tst() ' :L #AA# lskdf
Dim Act As Msig, Ept As Msig, Mthln
GoSub T1
Exit Sub
T1:
    Mthln = "Function MsigzL(Mthln) As Msig() '  ksdljf    #AA-BB#"
    Dim A0 As Arg: A0 = Arg(eByRefArgm, "Mthln", DftVty, "")
    Dim A() As Arg: PushArg A, A0
    Ept = Msig("", "Fun", "MsigzL", A, DftVty, "ksdljf", "#AA-BB#")
    GoTo Tst
Tst:
    Act = MsigzL(Mthln)
    If Not IsEqMsig(Act, Ept) Then Stop
    Return
End Sub
Function MsigzL(Mthln) As Msig
Dim L$: L = Mthln
Dim Mdy$: Mdy = ShfShtMdy(L):
Dim MTy$: MTy = ShfShtMthTy(L): If MTy = "" Then Thw CSub, "Given Mthln is invalid: No mth ty", "Mthln", Mthln
Dim Nm$: Nm = ShfNm(L)
Dim Chr$: Chr = ShfTyChr(L)
Dim Pm$: Pm = ShfBetBkt(L)
Dim Tyn$: Tyn = ShfNmAftAs(L)
Dim IsAy As Boolean: IsAy = ShfBkt(L)
Dim T As Vty: T = Vty(Chr, IsAy, Tyn)
Dim Rmk$: Rmk = RmvPfx(LTrim(L), "'")
Dim M$:  M = Memn(Rmk)
Rmk = Trim(Replace(Rmk, M, ""))
MsigzL = Msig(Mdy, MTy, Nm, ArgyzP(Pm), T, Rmk, M)
End Function

'Msigln**
Private Sub Msigln__Tst()
Dim Mthl$
GoSub T1
'GoSub T2
GoSub Z
Exit Sub
Z:
    BrwAy MsiglnyM
    Return
T1:
    Mthl = "Private Sub MsiglnyM__Tst()"
    Ept = "MsiglnyM__Tst_"
    GoTo Tst
T2:
    Mthl = "Function MsiglnyzM(M As CodeModule) As String()"
    Ept = "MsiglnyzM$() M:Md"
    GoTo Tst
Tst:
    Act = Msigln(MsigzL(Mthl))
    C
    'Debug.Print Act
    Return
End Sub
Private Sub MsiglnyM__Tst():                                 BrwAy MsiglnyM:         End Sub
Function MsiglnzL$(Mthln):                        MsiglnzL = Msigln(MsigzL(Mthln)):  End Function
Function MsiglnyM() As String():                  MsiglnyM = MsiglnyzM(CMd):         End Function
Function MsiglnyzM(M As CodeModule) As String(): MsiglnyzM = MsiglnyzL(MthlnyzM(M)): End Function
Function MsiglnyP() As String():                  MsiglnyP = MsiglnyzP(CPj):         End Function
Function MsiglnyzP(P As VBProject) As String():  MsiglnyzP = MsiglnyzL(MthlnyzP(P)): End Function
Function MsiglnyzL(Mthlny$()) As String()
Dim L: For Each L In Itr(Mthlny)
    PushI MsiglnyzL, MsiglnzL(L)
Next
End Function
Function Msigln$(S As Msig)
Dim P1$:   P1 = W1Part1(S)
Dim Pm$:   Pm = W1Pm(S.Arg)
Dim Rmk$: Rmk = AddPfxIfNB(S.ShtRmk, "' ")
       Msigln = JnSpcApNB(P1, Pm, S.Memn, Rmk)
End Function
Private Sub W1___Msigln(): End Sub
Private Function W1Part1$(S As Msig)
W1Part1 = S.Mthn & W1MthMdyChr(S.ShtMdy) & W1MthnSfx(S.ShtMdy, S.ShtTy, S.Ret)
End Function
Private Function W1MthMdyChr$(ShtMdy$)
Dim O$
Select Case True
Case ShtMdy = ""
Case ShtMdy = "Prv": O = "_"
Case ShtMdy = "Frd": O = "."
Case Else: Thw CSub, "ShtMdy error", "ShtMdy", ShtMdy
End Select
W1MthMdyChr = O
End Function
Private Function W1MthnSfx$(ShtMdy$, ShtTy$, RetVty As Vty)
If ShtTy = "Sub" Then Exit Function
W1MthnSfx = VarSfx(RetVty)
End Function
Private Function W1Pm$(A() As Arg)
Dim O$(), J%: For J = 0 To ArgUB(A)
    PushI O, ShtArg(A(J))
Next
W1Pm = QuoTerm(JnSpc(O))
End Function
