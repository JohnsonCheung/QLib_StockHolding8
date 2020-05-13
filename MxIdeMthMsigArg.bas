Attribute VB_Name = "MxIdeMthMsigArg"
Option Explicit
Option Compare Text
Const CNs$ = "Src.Mth"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMthMsigArg."
#If Doc Then
'Cml
' Arg Argument
' Argn Argument name
'Term
' ArgStr::S Comma term of method parameter
'
'Enum eArgM ! Argument modifier, which all string before Argn
'
' :    :S ! Argm Nm ArgSfx Dft
':ShtArgm: :S ! One-of-:ShtArgmAy
'            ! :Sfx: ::  TyChr[Bkt] | vbColon AsTy
'            ! :Dft: :: [ChrEq DftStr] !
':MthChr: :C #Mth-Ty-Chr# ! one of :TyChrLis
':TyChrLis: :S #Mth-Ty-Chr-List# ! !@#$%^&
':C:  :Chr #Char# ! One single char
':TyChr: :MthChr:
#End If
Public Const TyChrLis$ = "!@#$%^&"
Function Argn$(Arg)
Argn = TakNm(RmvArgm(Arg))
End Function

Function ShtArgzS$(ArgStr)
ShtArgzS = ShtArg(ArgzS(ArgStr))
End Function

Function ShtDft$(Dft$)
If Dft = "" Then Exit Function
If Left(Dft, 3) <> " = " Then Stop
ShtDft = "=" & Mid(Dft, 4)
End Function

Function ShtArgStryP() As String()
Dim Arg: For Each Arg In QSrt(AwDis(ArgStryP))
    PushI ShtArgStryP, ShtArgzS(Arg)
Next
End Function

Private Sub Argn__Tst()
Dim ArgStr$
'GoSub T1
GoSub YY
Exit Sub
T1:
    ArgStr = "Optional Fnn$"
    Ept = "Fnn"
    GoTo Tst
Tst:
    Act = Argn(ArgStr)
    C
    Return
YY:
Dim O() As S12
Dim A: For Each A In ArgStryzP(CPj)
    PushS12 O, S12(Argn(ArgStr), ArgStr)
Next
BrwS12y O
End Sub


Private Sub ShtArgStry__Tst()
Dim A() As S12
Dim Arg: For Each Arg In AwDis(ArgStryP)
    PushS12 A, S12(Arg, ShtArgzS(Arg))
Next
BrwS12y A
End Sub

Function ShtArgStry(ArgStry$()) As String()
Dim Arg: For Each Arg In Itr(ArgStry)
    PushI ShtArgStry, ShtArgzS(Arg)
Next
End Function
Function ShfMthTy$(OLin$)
Dim O$: O = TakMthTy(OLin$)
If O = "" Then Exit Function
ShfMthTy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function


Function IsTyChr(A) As Boolean
If Len(A) <> 1 Then Exit Function
IsTyChr = HasSubStr(TyChrLis, A)
End Function
Function TyChrzN$(Tyn)
Select Case Tyn
Case "Boolean":   TyChrzN = "*"
Case "String":   TyChrzN = "$"
Case "Integer":  TyChrzN = "%"
Case "Long":     TyChrzN = "&"
Case "Double":   TyChrzN = "#"
Case "Single":   TyChrzN = "!"
Case "Currency": TyChrzN = "@"
End Select
End Function

Function TynzTyChr$(TyChr$)
Const CSub$ = CMod & "TynzTyChr"
Dim O$
Select Case TyChr
Case "": O = "Variant"
Case "#": O = "Double"
Case "%": O = "Integer"
Case "!": O = "Signle"
Case "@": O = "Currency"
Case "^": O = "LongLong"
Case "$": O = "String"
Case "&": O = "Long"
Case Else: Thw CSub, "Invalid TyChr", "TyChr VdtTyChrLis", TyChr, TyChrLis
End Select
TynzTyChr = O
End Function

Function RmvTyChr$(S)
RmvTyChr = RmvLasChrzLis(S, TyChrLis)
End Function

Function Argm$(ArgStr)
Argm = PfxzAySpc(ArgStr, eArgmTxty)
End Function

Function RmvArgm$(ArgStr)
RmvArgm = ShfArgm(CStr(ArgStr))
End Function

Function ArgnyP() As String()
ArgnyP = ArgnyzP(CPj)
End Function

Function ArgnyzP(P As VBProject) As String()
Dim O$()
    Dim Mthln: For Each Mthln In MthlnyzP(P)
        PushIAy O, ArgStry(Mthln)
    Next
O = AwDis(O)
ArgnyzP = SrtAy(O)
End Function

Private Sub ShtArgzS__Tst()
Dim Arg$
GoSub Z
'GoSub T0
Exit Sub
Z:
    Dim S() As S12
    Dim A: For Each A In ArgStryP
        PushS12 S, S12(A, ShtArgzS(A))
    Next
    BrwS12y S
    Return
T0:
     Arg = "Optional UseVc As Boolean"
     Ept = "?UseVc?"
     GoTo Tst
Tst:
    Act = ShtArgzS(Arg)
    C
    Return
End Sub

Function ShtPm$(MthPm)
Dim O$()
Dim Arg: For Each Arg In Itr(SplitCommaSpc(MthPm))
    PushI O, ShtArgzS(Arg)
Next
ShtPm = JnSpc(O)
End Function

Function DclSfxzArg$(Arg$)
DclSfxzArg = DclSfx(ArgItm(Arg))
End Function

Function ShtDclSfxzArg$(Arg$)
ShtDclSfxzArg = ShtDclSfx(DclSfx(ArgItm(Arg)))
End Function

Function ArgItm$(Arg)
ArgItm = BefOrAll(RmvPfxSpc(RmvPfxSpc(Arg, "Optional"), "ParamArray"), " =")
End Function

Function FmtPm(Pm$, Optional IsNoBkt As Boolean) 'Pm is wo bkt.
Dim A$: A = Replace(Pm, "Optional ", "?")
Dim B$: B = Replace(A, " As ", ":")
Dim C$: C = Replace(B, "ParamArray ", "...")
If IsNoBkt Then
    FmtPm = C
Else
    FmtPm = QuoSq(C)
End If
End Function

Function RetAszDclSfx$(DclSfx)
If DclSfx = "" Then Exit Function
Dim B$
Dim F$: F = FstChr(DclSfx)
If IsTyChr(F) Then
    If Len(DclSfx) = 1 Then Exit Function
    B = RmvFstChr(DclSfx): If B <> "()" Then Stop
    RetAszDclSfx = " As " & TynzTyChr(F) & "()"
    Exit Function
End If
If TyChrzN(DclSfx) <> "" Then Exit Function
Select Case True
Case Left(DclSfx, 4) = " As ":   RetAszDclSfx = DclSfx
Case Left(DclSfx, 6) = "() As ": RetAszDclSfx = Mid(DclSfx, 3) & "()"
Case DclSfx = "()":              RetAszDclSfx = " As Variant()"
Case Else: Stop
End Select
End Function
Function TyChrzDclSfx$(DclSfx)
If Len(DclSfx) = 1 Then
    If IsTyChr(DclSfx) Then TyChrzDclSfx = DclSfx
End If
End Function

Function ArgSfxzRet$(Ret)
'Ret is either FunRetTyChr (in Sht-TyChr) or
'              FunRetAs    (The Ty-Str without As)
Select Case True
Case IsTyChr(FstChr(Ret)): ArgSfxzRet = Ret
Case HasSfx(Ret, "()") And TyChrzN(RmvSfx(Ret, "()")) <> "": ArgSfxzRet = TyChrzN(RmvSfx(Ret, "()")) & "()"
Case Else: ArgSfxzRet = " As " & Ret
End Select
End Function


Function ArgSfx$(Arg)
Const CSub$ = CMod & "ArgSfx"
Dim L$: L = Arg
ShfPfxSpc L, "Optional"
ShfPfxSpc L, "ByVal"
ShfPfxSpc L, "ParamArray"
If ShfNm(L) = "" Then Thw CSub, "Arg is invalid", "Arg", Arg
ArgSfx = ShfDclSfx(L)
End Function
Function ArgSfxy(Argy$()) As String()
Dim Arg: For Each Arg In Itr(Argy)
    PushI ArgSfxy, ArgSfx(Arg)
Next
End Function


Private Sub MthRetTy__Tst()
'Dim Mthln
'Dim A$:
'Mthln = "Function MthPm(MthPm$) As MthPm"
'A = MthRetTy(Mthln)
'Ass A.TyAsNm = "MthPm"
'Ass A.IsAy = False
'Ass A.TyChr = ""
'
'Mthln = "Function MthPm(MthPm$) As MthPm()"
'A = MthRetTy(Mthln)
'Ass A.TyAsNm = "MthPm"
'Ass A.IsAy = True
'Ass A.TyChr = ""
'
'Mthln = "Function MthPm$(MthPm$)"
'A = MthRetTy(Mthln)
'Ass A.TyAsNm = ""
'Ass A.IsAy = False
'Ass A.TyChr = "$"
'
'Mthln = "Function MthPm(MthPm$)"
'A = MthRetTy(Mthln)
'Ass A.TyAsNm = ""
'Ass A.IsAy = False
'Ass A.TyChr = ""
End Sub


Function ArgStry(Mthln) As String()
ArgStry = SplitCommaSpc(MthPm(Mthln))
End Function

Function ArgStryzMthPmAy(MthPmAy$()) As String()
Dim MthPm: For Each MthPm In Itr(MthPmAy)
    PushIAy ArgStryzMthPmAy, SplitCommaSpc(MthPm)
Next
End Function

Function ArgStryzL(Mthlny$()) As String()
Dim Mthln: For Each Mthln In Itr(Mthlny)
    PushIAy ArgStryzL, ArgStry(Mthln)
Next
End Function
'-----
Private Sub ArgStryP__Tst()
VcAy ArgStryP
End Sub

Function ArgStryP() As String()
ArgStryP = ArgStryzP(CPj)
End Function
'---
Function ArgStryzP(P As VBProject) As String()
ArgStryzP = ArgStryzL(MthlnyzP(P))
End Function

Function ArgStryzPmAy(PmAy$()) As String()
Dim Pm, Arg
For Each Pm In Itr(PmAy)
    For Each Arg In Itr(SplitCommaSpc(Pm))
        PushI ArgStryzPmAy, Arg
    Next
Next
End Function

Function NArg(Mthln) As Byte
NArg = Si(SplitComma(BetBkt(Mthln)))
End Function
Function ArgNy(Arg() As Arg) As String()
Dim J%: For J = 0 To ArgUB(Arg)
    PushI ArgNy, Arg(J).Argm
Next
End Function

Function ArgNyzPm(Pm$) As String()
Dim Ay$(): Ay = Split(Pm, ", ")
Dim I
For Each I In Itr(Ay)
    PushI ArgNyzPm, TakNm(I)
Next
End Function

Function ArgyzP(Pm$) As Arg() ' @Pm is the Bet-Bkt-of a Mthln
Dim A: For Each A In Itr(SplitCommaSpc(Pm))
    PushArg ArgyzP, ArgzS(A)
Next
End Function
'--
Sub ArgzS__Tst()
Dim O$()
Dim S: For Each S In ArgStryP
    PushI O, ShtArg(ArgzS(S))
Next
VcAy QSrt(AwDis(O))
End Sub
Function ArgzS(ArgStr) As Arg
Dim S$:                 S = Trim(ArgStr)
Dim Mdy As eArgm:     Mdy = ShfArgm(S)
Dim Nm$:               Nm = ShfNm(S): If Nm = "" Then Thw CSub, "Invalid ArgStr: no name", "ArgStr", ArgStr
Dim Chr$:             Chr = ShfTyChr(S)
Dim Tyn$:             Tyn = SzTrue(Chr = "", ShfNmAftAs(S))
Dim IsAy As Boolean: IsAy = ShfBkt(S)
Dim Dft$:             Dft = SzTrue(ShfTermX(S, "="), S)
Dim Ty As Vty:         Ty = Vty(Chr, IsAy, Tyn)
                    ArgzS = Arg(Mdy, Nm, Ty, Dft)
End Function

Function eArgm(S) As eArgm
If S = "" Then Exit Function
eArgm = IxzAy(eArgmTxty, S)
End Function

Function eArgmSy() As String()
Static X$(): If Si(X) = 0 Then X = Termy(eArgmSS)
eArgmSy = X
End Function
Function ShfArgm(OArg$) As eArgm
ShfArgm = eArgm(ShfPfxyS(OArg, eArgmTxty, "ByRef"))
End Function

Function ShfArgSfx$(OLin$)
Dim P%: P = InStr(OLin, "=")
If P > 0 Then
    ShfArgSfx = Left(OLin, P - 2)
    OLin = Mid(OLin, P - 1)
    Exit Function
Else
    ShfArgSfx = OLin
    OLin = ""
End If
End Function

'---
Function ShtArg$(A As Arg) ' Will have no spc in ShtArg
Dim Pfx$: Pfx = W1ShtArgm(A.Argm)
Dim Sfx$: Sfx = VarSfx(A.Ty)
Dim Dft$: Dft = SzTrue(A.Dft <> "", "=" & A.Dft)
       ShtArg = Pfx & A.Argn & Sfx & Dft
End Function
Private Function W1ShtArgm$(M As eArgm)
Dim O$
Select Case True
Case M = eByRefArgm: O = ""
Case M = eByValArgm: O = "*"
Case M = eOptByRefArgm: O = "?"
Case M = eOptByValArgm: O = "?*"
Case M = ePmArgm: O = ".."
Case Else: EnmEr CSub, "eArgm", eArgmSS, M
End Select
W1ShtArgm = O
End Function
