Attribute VB_Name = "MxVbStrLines"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str.Lines"
Const CMod$ = CLib & "MxVbStrLines."

Function AddLines$(A$, B$)
If LTrim(A) = "" Then AddLines = B: Exit Function
If LTrim(B) = "" Then AddLines = A: Exit Function
AddLines = A & vbCrLf & B
End Function

Function AddLinesAp$(ParamArray LinesAp())
Dim Av(): Av = LinesAp
AddLinesAp = AddLinesAv(Av)
End Function

Function AddLinesAv$(LinesAv())
If Si(LinesAv) = 0 Then Exit Function
Dim O$(): O = AwNB(LinesAv)
AddLinesAv = JnCrLf(O)
End Function

Function WdtzLines%(Lines)
WdtzLines = WdtzAy(SplitCrLf(Lines))
End Function

Function WdtzLinesy%(Linesy$())
Dim O%
Dim Lines: For Each Lines In Itr(Linesy)
    O = Max(O, WdtzLines(Lines))
Next
WdtzLinesy = O
End Function

Sub VcLinesy(Linesy$())
Vc FmtLinesy(Linesy)
End Sub

Sub BrwLinesy(Linesy$())
B FmtLinesy(Linesy)
End Sub

Private Sub FmtLinesy__Tst()
Dim Linesy
GoSub Z
Exit Sub
Z:
    BrwLinesy SampLinesy
    Return
End Sub

Function SampLinesy() As String()
BfrClr
BfrV RplVbl("sdklf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf")
BfrV RplVbl("sdklf2-49230  sdfjldf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf")
BfrV RplVbl("sdsdfklf2-49230  sdfjldf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf")
SampLinesy = BfrLy
End Function

Function FmtLinesy(Linesy$(), Optional BegIx%) As String()
If Si(Linesy) = 0 Then Exit Function
Dim W%: W = WdtzLinesy(Linesy)
Dim VbarDashSepLn: VbarDashSepLn = Quo(Dup("-", W + 2), "|")
Dim Lines
PushI FmtLinesy, VbarDashSepLn
For Each Lines In Itr(Linesy)
    PushIAy FmtLinesy, AddIxPfxzLineszW(Lines, W, BegIx)
    PushI FmtLinesy, VbarDashSepLn
Next
End Function
Function AddIxPfxzLineszW(Lines, W%, Optional BegIx%) As String()
Dim L
For Each L In Itr(SplitCrLf(Lines))
    PushI AddIxPfxzLineszW, "| " & AliL(L, W) & " |"
Next
End Function

Function EnsBet(I, A, B)
Select Case True
Case I < A: EnsBet = A
Case I > B: EnsBet = B
Case Else: EnsBet = I
End Select
End Function


Private Sub RTrimLines__Tst()
Dim Lines$: Lines = LineszVbl("lksdf|lsdfj|||")
Dim Act$: Act = RTrimLines(Lines)
Debug.Print Act & "<"
Stop
End Sub

Private Sub LineszLasN__Tst()
Dim Ay$(), A$, J%
For J = 0 To 9
Push Ay, "Line " & J
Next
A = Join(Ay, vbCrLf)
Debug.Print LineszLasN(A, 3)
End Sub

Function FstLin(Lines$)
FstLin = BefOrAll(Lines, vbCrLf)
End Function

Function LinesApp$(A, L)
If A = "" Then LinesApp = L: Exit Function
LinesApp = A & vbCrLf & L
End Function

Private Sub RTrimLines1__Tst()
Dim Lines$: Lines = LineszVbl("lksdf|lsdfj|||")
Dim Act$: Act = RTrimLines(Lines)
Debug.Print Act & "<"
Stop
End Sub

Function LineszLasN$(Lines$, N%)
LineszLasN = JnCrLf(AwLasN(SplitCrLf(Lines), N))
End Function

Function LnCnt&(Lines)
LnCnt = Si(SplitCrLf(Lines))
End Function

Function MaxLnCnt&(Linesy$())
Dim O&, Lines: For Each Lines In Itr(Linesy)
    O = Max(O, LnCnt(Lines))
Next
MaxLnCnt = O
End Function

Function SqhzLines(Lines$) As Variant()
SqhzLines = Sqh(SplitCrLf(Lines))
End Function

Function SqvzLines(Lines$) As Variant()
SqvzLines = Sqv(SplitCrLf(Lines))
End Function

Function RTrimLines$(Lines)
Dim At&
For At = Len(Lines) To 1 Step -1
    If Not IsStrAtSpcCrLf(Lines, At) Then RTrimLines = Left(Lines, At): Exit Function
Next
End Function

Function LasLin(Lines$)
LasLin = LasEle(SplitCrLf(Lines))
End Function

Function LineszAli$(Lines$, W%)
Const CSub$ = CMod & "LineszAli"
Dim Las$: Las = LasLin(Lines)
Dim N%: N = W - Len(Las)
If N > 0 Then
    LineszAli = Lines & Space(N)
Else
    Warn CSub, "W is too small", "Lines.LasLin W", Las, W
    LineszAli = Lines
End If
End Function

Function NLn&(Lines)
Dim R As RegExp: Set R = Rx("\n", MultiLine:=True, IsGlobal:=True)
NLn = CntzRx(Lines, R)
End Function

Function LineszV$(V)
LineszV = JnCrLf(FmtV(V))
End Function



Function AddIxPfxzLines(Lines, Optional BegIx%) As String()
AddIxPfxzLines = AmAddIxPfx(SplitCrLf(Lines), BegIx)
End Function
