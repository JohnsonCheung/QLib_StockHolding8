Attribute VB_Name = "MxVbStrSub"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str"
Const CMod$ = CLib & "MxVbStrSub."

Function LasChr$(S)
LasChr = Right(S, 1)
End Function

Function Las2Chr$(S)
Las2Chr = Right(S, 2)
End Function

Function SndChr$(S)
SndChr = Mid(S, 2, 1)
End Function

Function FstAsc%(S)
FstAsc = Asc(FstChr(S))
End Function
Function SndAsc%(S)
SndAsc = Asc(SndChr(S))
End Function
Function UCasFst$(S)
UCasFst = UCase(FstChr(S)) & RmvFstChr(S)
End Function
Function FstChr$(S)
FstChr = Left(S, 1)
End Function

Function Fst2Chr$(S)
Fst2Chr = Left(S, 2)
End Function

Function CntSubStr&(S, SubStr$)
Dim P&: P = 1
Dim O&, L%
L = Len(SubStr)
While P > 0
    P = InStr(P, S, SubStr)
    If P = 0 Then CntSubStr = O: Exit Function
    O = O + 1
    P = P + L
Wend
End Function

Private Sub CntSubStr__Tst()
Dim A$, SubStr$
A = "aaaa":                 SubStr = "aa":  Ept = CLng(2): GoSub Tst
A = "aaaa":                 SubStr = "a":   Ept = CLng(4): GoSub Tst
A = "skfdj skldfskldf df ": SubStr = " ":   Ept = CLng(3): GoSub Tst
Exit Sub
Tst:
    Act = CntSubStr(A, SubStr)
    C
    Return
End Sub

Function CntDot&(S)
CntDot = CntSubStr(S, ".")
End Function
