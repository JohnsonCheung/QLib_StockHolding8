Attribute VB_Name = "MxVbStrLnWrp"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrLnWrp."

'**Wrp-Fun
Function ShdWrp(Str$, Optional Wdt% = 80) As Boolean
Select Case True
Case IsLines(Str), _
    Len(Str) > Wdt
    ShdWrp = True
End Select
End Function

'**Fm-Ly
Function LnzLy$(Ly$()): LnzLy = JnSpc(AmRTrim(Ly)): End Function

'**Shf-Ln
Function ShfLn$(OLin$, W%)
If OLin = "" Then Exit Function
Dim N%: N = W1NLeft(OLin$, W%)
ShfLn = Left(OLin, N)
OLin = LTrim(Mid(OLin, N + 1))
End Function
Private Function W1NLeft%(Ln$, W%)
Dim L$: L = Left(Ln, W)
Dim P&: P = InStrRev(L, " ")
Select Case True
Case Len(L)
    W1NLeft = Len(L)
Case LasChr(L) = " "
    W1NLeft = Len(RTrim(L))
Case P = 0
    W1NLeft = Len(L)
Case Else
    W1NLeft = P - 1
End Select
End Function

'**Wrp-Ln-Plus
Private Sub WrpLy__Tst()
Dim Ly$(), Wdt%
GoSub T1
Exit Sub
T1:
    Ly = Sy("a b c d")
    Wdt = 80
    Ept = Sy("a b c d")
    GoTo Tst
Tst:
    Act = WrpLy(Ly, Wdt)
    C
    Return
End Sub
Private Sub WrpLn__Tst()
Dim W%, A$
GoSub Z1
'GoSub Z2
Exit Sub
Z1:
    W = 80
    A = "AddColMthl DoCachedMthcP DoCachedMthczP MthDrs MthcDrs MthcDrsM MthcDrsP MthcDrsP__Tst MthcDrszFxa MthcDrszM MthcDrszP MthcDrszPjf MthcDrszPjfy MthcDrszV MthDrsM MthDrsP MthDrszM MthDrszP PFunDrszP MthDr MthlnDr WsoMthcP Z_MthcDrszP"
    GoTo Tst
Z2:
    A = "lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf sdklf sdklfj dsfj "
    W = 80
    Ept = Sy("lksjf lksdj flksdjf lskdjf lskdjf lksdjf lksdjf klsdjf klj skldfj lskdjf klsdjf ", _
    "klsdfj klsdfj lskdfj  sdlkfj lsdkfj lsdkjf klsdfj lskdjf lskdjf kldsfj lskdjf", _
    "sdklf sdklfj dsfj ")
    GoTo Tst
Z3:
    A = "DymJnDot Ln X Var Box ULin Brw Ly Lines IEr_Er XTab CrtTblUSysRegInf EnsTblUSysRegInf"
    W = 80
    GoTo Tst
Tst:
    Act = WrpLn(A, W)
    'C
    Dmp Act
    Return
End Sub
Function WrpStr$(Str$, Optional Wdt% = 80): WrpStr = JnCrLf(WrpLy(SplitCrLf(Str), Wdt)): End Function
Function WrpLy(Ly$(), Optional Wdt% = 80) As String()
Dim Ln: For Each Ln In Itr(Ly)
    PushIAy WrpLy, WrpLn(Ln, Wdt)
Next
End Function
Function WrpLnzW(Ln, Optional W% = 80) As String()
Dim L$: L = RTrim(Ln)
PushI WrpLnzW, ShfLn(L, W)
Dim J%: While L <> ""
    LoopTooMuch CSub, J
    PushI WrpLnzW, ShfLn(L, W)
Wend
End Function
Function WrpLn(Ln, Optional W% = 80, Optional SndLnNSpc% = 4, Optional SndLnPfx$ = ". ", Optional Ln1Lbl$) As String() '@W is inl Ln1-Lbl
Dim L$: L = RTrim(Ln)
PushIAy WrpLn, W1ShfLn1()
Dim J%: While L <> ""
    LoopTooMuch CSub, J
    PushIAy WrpLn, W1ShfRstLn(L, W)
Wend
End Function
Private Function W1ShfLn1() As String()
End Function
Private Function W1ShfRstLn(L$, W%) As String()
End Function
