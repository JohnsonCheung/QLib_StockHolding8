Attribute VB_Name = "MxIdeMthMsigUd"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeMthMsigUdt."
#If Doc Then
#End If
Enum eArgm: eByRefArgm: eByValArgm: eOptByValArgm: eOptByRefArgm: ePmArgm: End Enum ' #Arguement-Modifier#
Public Const eArgmSS$ = "ByRefArgm ByValArgm OptByValArgm OptByRefArgm PmArgm" ' Dot means

Type Vty: TyChr As String: IsAy As Boolean: Tyn As String: End Type ' Deriving(Ctor Ay) #Variable-Type# OptMbr(Tyn IsAy)
Type Arg: Argm As eArgm: Argn As String: Ty As Vty: Dft As String: End Type 'Deriving(Ay Ctor) OptMbr(Dft)
Type Msig
    ShtMdy As String
    ShtTy As String  ' ShtMthTy
    Mthn As String
    Arg() As Arg
    Ret As Vty
    ShtRmk As String ' On the same lines
    Memn As String
End Type ' Deriving(Ay Ctor)
Function IsEqMsig(A As Msig, B As Msig) As Boolean
With A
    Select Case True
    Case .Mthn <> B.Mthn
    Case .ShtMdy <> B.ShtMdy
    Case .ShtRmk <> B.ShtRmk
    Case .ShtTy <> B.ShtTy
    Case IsEqArgy(.Arg, B.Arg)
    Case Else: IsEqMsig = True
    End Select
End With
End Function
Function IsEqArgy(A() As Arg, B() As Arg) As Boolean
Dim U&: U = ArgUB(A): If U <> ArgUB(B) Then Exit Function
Dim J%: For J = 0 To U
    If IsEqArg(A(J), B(J)) Then Exit Function
Next
IsEqArgy = True
End Function
Function IsEqArg(A As Arg, B As Arg) As Boolean
With A
    Select Case True
    Case .Argm <> B.Argm
    Case .Argn <> B.Argn
    Case .Dft <> B.Dft
    Case Not IsEqVty(.Ty, B.Ty)
    Case Else: IsEqArg = True
    End Select
End With
End Function
Function IsEqVty(A As Vty, B As Vty) As Boolean
With A
    Select Case True
    Case .IsAy <> B.IsAy
    Case .TyChr <> B.TyChr
    Case .Tyn <> B.Tyn
    Case Else: IsEqVty = True
    End Select
End With
End Function

Function VarSfx$(T As Vty) ' #Variable-Sfx# short variable suffix directly added to varn which can be a dimn argn udtn
Dim Bkt$: Bkt = BktzIsAy(T.IsAy)
Select Case True
Case T.Tyn = "" And T.TyChr = "": VarSfx = Bkt
Case T.Tyn = "":                  VarSfx = T.TyChr & Bkt
Case Else
    Dim C$: C = TyChrzN(T.Tyn)
    If C = "" Then
        VarSfx = ":" & ShtTyn(T.Tyn) & Bkt
    Else
        VarSfx = C & Bkt
    End If
End Select
End Function

Function eArgmTxty() As String()
Const T0 = "ByRef"
Const T1 = "ByVal"
Const T2 = "Optional ByVal"
Const T3 = "Optional ByRef"
Const T4 = "ParamArray"
Static X$(): If Si(X) = 0 Then X = Sy(T0, T1, T2, T3, T4)
eArgmTxty = X
End Function

Sub PushArgStry(O() As Arg, A() As Arg): Dim J&: For J = 0 To ArgUB(A): PushArg O, A(J): Next: End Sub

Private Function W1ShtArgSfxzS$(ArgSfx$)
Dim O$
Select Case True
Case ArgSfx = ""
Case ArgSfx = "()":             O = "()"
Case ArgSfx = "() As Boolean":  O = "~()"
Case ArgSfx = " As Boolean":    O = "~"
Case Else
    Dim A$: A = ArgSfx
    Dim IsAy As Boolean
    IsAy = ShfPfx(A, "()")
    If Not IsAy Then IsAy = ShfSfx(A, "()")
    If ShfPfx(A, " As ") Then
        O = ":" & ShtTyn(A) & IIf(IsAy, "()", "")
        Exit Function
    End If
    '---
    If Len(A) <> 1 Then Stop
    If Not HasSubStr(TyChrLis, A) Then Stop
    O = ArgSfx
End Select
W1ShtArgSfxzS = O
End Function
Function DftVty() As Vty: DftVty = Vty("", False, ""): End Function
Function Vty(TyChr, IsAy, Tyn) As Vty
With Vty
    .TyChr = TyChr
    .IsAy = IsAy
    .Tyn = Tyn
End With
End Function
Function AddVty(A As Vty, B As Vty) As Vty(): PushVty AddVty, A: PushVty AddVty, B: End Function
Sub PushVtyAy(O() As Vty, A() As Vty): Dim J&: For J = 0 To VtyUB(A): PushVty O, A(J): Next: End Sub
Sub PushVty(O() As Vty, M As Vty): Dim N&: N = VtySi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function VtySi&(A() As Vty): On Error Resume Next: VtySi = UBound(A) + 1: End Function
Function VtyUB&(A() As Vty): VtyUB = VtySi(A) - 1: End Function
Function Arg(Argm As eArgm, Argn, Ty As Vty, Dft) As Arg
With Arg
    .Argm = Argm
    .Argn = Argn
    .Ty = Ty
    .Dft = Dft
End With
End Function
Function AddArg(A As Arg, B As Arg) As Arg(): PushArg AddArg, A: PushArg AddArg, B: End Function
Sub PushArgAy(O() As Arg, A() As Arg): Dim J&: For J = 0 To ArgUB(A): PushArg O, A(J): Next: End Sub
Sub PushArg(O() As Arg, M As Arg): Dim N&: N = ArgSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function ArgSi&(A() As Arg): On Error Resume Next: ArgSi = UBound(A) + 1: End Function
Function ArgUB&(A() As Arg): ArgUB = ArgSi(A) - 1: End Function
Function Msig(ShtMdy, ShtTy, Mthn, Arg() As Arg, Ret As Vty, ShtRmk, Memn) As Msig
With Msig
    .ShtMdy = ShtMdy
    .ShtTy = ShtTy
    .Mthn = Mthn
    .Arg = Arg
    .Ret = Ret
    .ShtRmk = ShtRmk
    .Memn = Memn
End With
End Function
Function AddMsig(A As Msig, B As Msig) As Msig(): PushMsig AddMsig, A: PushMsig AddMsig, B: End Function
Sub PushMsigAy(O() As Msig, A() As Msig): Dim J&: For J = 0 To MsigUB(A): PushMsig O, A(J): Next: End Sub
Sub PushMsig(O() As Msig, M As Msig): Dim N&: N = MsigSi(O): ReDim Preserve O(N): O(N) = M: End Sub
Function MsigSi&(A() As Msig): On Error Resume Next: MsigSi = UBound(A) + 1: End Function
Function MsigUB&(A() As Msig): MsigUB = MsigSi(A) - 1: End Function


