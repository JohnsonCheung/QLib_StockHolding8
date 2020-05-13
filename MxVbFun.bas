Attribute VB_Name = "MxVbFun"
Option Compare Text
Option Explicit
Public Const CLib$ = "?."
Public Const CMod$ = CLib & "MxFun."
Public Const CSub$ = CMod & "?"

Sub Swap(OA, OB)
Dim X: X = OA
OA = OB
OB = X
End Sub

Function CUsr$()
CUsr = Environ$("USERNAME")
End Function

Sub Asg(Fm, _
Into)
Select Case True
Case IsObject(Fm): Set Into = Fm
Case Else:             Into = Fm
End Select
End Sub

Sub ShellHid(Fcmd$)
Shell Fcmd, vbHide
End Sub

Sub ShellMax(Fcmd$)
Shell Fcmd, vbMaximizedFocus
End Sub

Sub ChkBet(Fun$, V, FmV, ToV)
If FmV > V Then Thw Fun, "FmV > V", "V FmV ToV", V, FmV, ToV
If ToV < V Then Thw Fun, "ToV < V", "V FmV ToV", V, FmV, ToV
End Sub

Function InStrWiIthSubStr&(S, SubStr, Optional Ith% = 1)
Const CSub$ = CMod & "InStrWiIthSubStr"
Dim P&, J%
If Ith < 1 Then Thw CSub, "Ith cannot be <1", "Ith", Ith
For J = 1 To Ith
    P = InStr(P + 1, SubStr, SubStr)
    If P = 0 Then Exit Function
Next
InStrWiIthSubStr = P
End Function

Function InStrN&(S, SubStr, Optional N% = 1)
InStrN = InStrWiIthSubStr(S, SubStr, N)
End Function

Function CvNothing(A)
If IsEmpty(A) Then Set CvNothing = Nothing: Exit Function
Set CvNothing = A
End Function

Private Sub InStrN__Tst()
Dim Act&, Exp&, S, SubStr, N%

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 1
Exp = 1
Act = InStrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 2
Exp = 6
Act = InStrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 3
Exp = 11
Act = InStrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 4
Exp = 0
Act = InStrN(S, SubStr, N)
Ass Exp = Act
End Sub

Function Max(A, B, ParamArray Ap())
Dim O: O = IIf(A > B, A, B)
Dim J%: For J = 0 To UBound(Ap)
   If Ap(J) > O Then O = Ap(J)
Next
Max = O
End Function

Function CvVbTy(A) As VbVarType
CvVbTy = A
End Function

Function CprMth(C As eCas) As VbCompareMethod
Select Case True
Case C = eCasSen: CprMth = vbBinaryCompare
Case C = eIgnCas: CprMth = vbTextCompare
Case Else: Thw CSub, "Invalid value of eCas", "eCas", C
End Select
End Function
Function CanCvLng(V) As Boolean
On Error GoTo X
Dim L&: L = CLng(V)
CanCvLng = True
X:
End Function

Function MinUB(Ay1, Ay2)
MinUB = Min(UB(Ay1), UB(Ay2))
End Function

Function Min(ParamArray A())
Dim O, J&, Av()
Av = A
Min = MinEle(Av)
End Function

Function MaxVbTy(A As VbVarType, B As VbVarType) As VbVarType
Dim O As VbVarType
If A = vbString Or B = vbString Then O = A: Exit Function
If A = vbEmpty Then O = B: Exit Function
If B = vbEmpty Then O = A: Exit Function
If A = B Then O = A: Exit Function
Dim AIsNum As Boolean, BIsNum As Boolean
AIsNum = IsVbTyNum(A)
BIsNum = IsVbTyNum(B)
Select Case True
Case A = vbBoolean And BIsNum: O = B
Case AIsNum And B = vbBoolean: O = A
Case A = vbDate Or B = vbDate: O = vbString
Case AIsNum And BIsNum:
    Select Case True
    Case A = vbByte: O = B
    Case B = vbByte: O = A
    Case A = vbInteger: O = B
    Case B = vbInteger: O = A
    Case A = vbLong: O = B
    Case B = vbLong: O = A
    Case A = vbSingle: O = B
    Case B = vbSingle: O = A
    Case A = vbDouble: O = B
    Case B = vbDouble: O = A
    Case A = vbCurrency Or B = vbCurrency: O = A
    Case Else: Stop
    End Select
Case Else: Stop
End Select
End Function

Sub SndKeys(A$)
DoEvents
SendKeys A, True
End Sub

Function NDig&(N)
Dim A$: A = N
NDig = Len(A)
End Function



Function SumLngAy@(A&())
Dim L, O@
For Each L In Itr(A)
    O = O + L
Next
SumLngAy = O
End Function

Private Sub SumLngAy__Tst()
Dim S$: S = LineszFt(PjfP)
Debug.Assert SumLngAy(AscCntAy(S)) = Len(S)
End Sub

Function AscCntAy(S) As Long()
Dim O&(255), A As Byte
Dim J&
For J = 1 To Len(S)
    A = Asc(Mid(S, J, 1))
    O(A) = O(A) + 1
Next
AscCntAy = O
End Function
Function RemBlkSz&(N&, BlkSz%)
RemBlkSz = (N Mod BlkSz)
End Function

Function NBlk&(N&, BlkSz%)
NBlk = ((N - 1) \ BlkSz) + 1
End Function
