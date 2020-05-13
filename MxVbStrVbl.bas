Attribute VB_Name = "MxVbStrVbl"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrVbl."

Function VblzLines$(Lines$)
VblzLines = Replace(RmvCr(Lines), vbLf, "|")
End Function

Function DiczVbl(Vbl$, Optional JnSep$ = vbCrLf) As Dictionary
Set DiczVbl = Dic(SplitVBar(Vbl), JnSep)
End Function

Function SyzVbl(Vbl) As String()
SyzVbl = SplitVBar(Vbl)
End Function
Function ItrzVbl(Vbl)
ItrzVbl = Itr(SyzVbl(Vbl))
End Function

Function LineszVbl$(Vbl)
LineszVbl = Replace(Vbl, "|", vbCrLf)
End Function

Function IsVbl(S) As Boolean
Select Case True
Case Not IsStr(S)
Case HasSubStr(S, vbCr)
Case HasSubStr(S, vbLf)
Case Else: IsVbl = True
End Select
End Function

Function IsVblAy(VblAy$()) As Boolean
Dim Vbl: For Each Vbl In Itr(VblAy)
    If Not IsVbl(Vbl) Then Exit Function
Next
IsVblAy = True
End Function

Function IsVdtVbl(Vbl$) As Boolean
If HasSubStr(Vbl, vbCr) Then Exit Function
If HasSubStr(Vbl, vbLf) Then Exit Function
IsVdtVbl = True
End Function

Function DrszTRst(FF$, TRstLy$()) As Drs
DrszTRst = DrszFF(FF, DyoTRst(TRstLy))
End Function
Function DyoTRst(TRstLy$()) As Variant()
Dim L: For Each L In Itr(TRstLy)
    PushI DyoTRst, SyzTRst(L)
Next
End Function
Function DyoTLny(TLny$()) As Variant()
Dim I
For Each I In Itr(TLny)
    PushI DyoTLny, Termy(I)
Next
End Function

Function DyoVblLy(A$()) As Variant()
Dim I
For Each I In Itr(A)
    PushI DyoVblLy, AmTrim(SplitVBar(I))
Next
End Function
Function DyoSSVbl(SSVbl$) As Variant()
Dim SS: For Each SS In Itr(SplitVBar(SSVbl))
    PushI DyoSSVbl, SyzSS(SS)
Next
End Function

Private Sub DyoVblLy__Tst()
Dim VblLy$()
GoSub T1
Exit Sub
T0:
    BfrClr
    BfrV "1 | 2 | 3"
    BfrV "4 | 5 6 |"
    BfrV "| 7 | 8 | 9 | 10 | 11 |"
    BfrV "12"
    VblLy = BfrLy
    Ept = Array(SyzSS("1 2 3"), Sy("4", "5 6", ""), Sy("", "7", "8", "9", "10", "11", ""), Sy("12"))
    GoTo Tst
Exit Sub
T1:
    BfrClr
    BfrV "|lskdf|sdlf|lsdkf"
    BfrV "|lsdf|"
    BfrV "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
    BfrV "|sdf"
    VblLy = BfrLy
    Ept = ""
    GoTo Tst
Tst:
    Act = DyoVblLy(VblLy)
    DmpDy CvAv(Act)
'    C
    Return
End Sub

Function LyzVbl(Vbl) As String()
LyzVbl = SplitVBar(Vbl)
End Function
