Attribute VB_Name = "MxVbStrHas"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrHas."
Public Const NNoKw$ = "Option Compare Text"
Enum eCas: eIgnCas: eCasSen: End Enum

Sub XX1()
#If False Then
As
Boolean
Const
Dim
Each
Else
Empty
End
Exit
Explicit
False
For
Function
Get
If
In
Me
New
Next
Not
Optional
Private
Property
Set
Sub
Then
True
Variant
#End If
End Sub

Function HasDot(S) As Boolean
HasDot = HasSubStr(S, ".")
End Function

Function HasSngQ(S) As Boolean
HasSngQ = InStr(S, vbSngQ)
End Function

Function HasDblQ(S) As Boolean
HasDblQ = InStr(S, vbDblQ)
End Function

Private Sub RmvBetDblQ__Tst()
Dim S$
GoSub T1
GoSub T2
GoSub YY
Exit Sub
T1:
    S = "("""""""")"
    Ept = "("""""""")"
    GoTo Tst
T2:
    S = "For Each I In AwSubStr(AwSubStr(SrczP(CPj), ""'"")"
    Ept = "For Each I In AwSubStr(AwSubStr(SrczP(CPj), """")"
    GoTo Tst
Tst:
    Act = RmvBetDblQ(S)
    C
    Return
YY:
    Dim S12y() As S12
    Dim L: For Each L In SrczP(CPj)
        If Not IsRmkln(L) Then
            If HasDblQ(L) Then
                PushS12 S12y, S12(L, RmvBetDblQ(L))
            End If
        End If
    Next
    BrwS12y S12y
    Return
End Sub

Function RmvBetDblQ$(S)
Dim P&: P = InStr(S, vbDblQ)
Dim O$: O = S
While P > 0
    Dim J%: J = J + 1: If J > 10000 Then Stop
    Dim P1&: P1 = InStr(P + 1, O, vbDblQ): If P1 = 0 Then Stop
    O = Left(O, P) & Mid(O, P1)
    P = InStr(P + 2, O, vbDblQ)
Wend
RmvBetDblQ = O
End Function

Function NoLf(S) As Boolean
NoLf = Not HasLf(S)
End Function

Function HasLf(S) As Boolean
HasLf = HasSubStr(S, vbLf)
End Function

Function DblQCnt%(S): DblQCnt = SubStrCnt(S, vbDblQ): End Function

Function SubStrCnt%(S, SubStr, Optional C As eCas = eIgnCas)
Dim L%: L = Len(SubStr)
Dim M%, J%, P%, O%
P = 1
Again:
    LoopTooMuch CSub, J
    M = InStr(P, S, SubStr): If M = 0 Then SubStrCnt = O: Exit Function
    P = P + M + L
    O = O + 1
    GoTo Again
End Function

Function HasSubStr(S, SubStr, Optional C As eCas = eIgnCas, Optional FmPos% = 1) As Boolean
HasSubStr = InStr(FmPos, S, SubStr, CprMth(C)) > 0
End Function

Function HasCrLf(S) As Boolean
HasCrLf = HasSubStr(S, vbCrLf)
End Function

Function HasHyphen(S) As Boolean
HasHyphen = HasSubStr(S, "-")
End Function

Function HasPound(S) As Boolean
HasPound = InStr(S, "#") > 0
End Function

Function HasSpc(S) As Boolean
HasSpc = InStr(S, " ") > 0
End Function

Function HasSqBkt(S) As Boolean
HasSqBkt = FstChr(S) = "[" And LasChr(S) = "]"
End Function

Function HasChrList(S, ChrList$, Optional Cpr As VbCompareMethod) As Boolean
Dim J%
For J = 1 To Len(ChrList)
    If HasSubStr(S, Mid(ChrList, J, 1), Cpr) Then HasChrList = True: Exit Function
Next
End Function

Function HasSubStrAy(S, SubStrAy$()) As Boolean
Dim SubStr
For Each SubStr In SubStrAy
    If HasSubStr(S, SubStr) Then HasSubStrAy = True: Exit Function
Next
End Function
Function HasTT(S, T1, T2) As Boolean
HasTT = Has2T(S, T1, T2)
End Function

Function HasVbar(S) As Boolean
HasVbar = HasSubStr(S, "|")
End Function
