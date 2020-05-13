Attribute VB_Name = "MxVbFmtAliLy"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CNs$ = "Str.Ly"
Const CMod$ = CLib & "MxVbFmtAliLy."
Enum eAli: eAliLeft: eAliRight: End Enum

Function AliLyz1T(Ly$()) As String(): AliLyz1T = AliLyzNTerm(Ly, 1): End Function
Function AliLyz2T(Ly$()) As String(): AliLyz2T = AliLyzNTerm(Ly, 2): End Function
Function AliLyz3T(Ly$()) As String(): AliLyz3T = AliLyzNTerm(Ly, 3): End Function
Function AliLyz4T(Ly$()) As String(): AliLyz4T = AliLyzNTerm(Ly, 4): End Function

Function AliLyzNTerm(Ly$(), N%) As String(): AliLyzNTerm = AmRTrim(FmtStrColy(NTermRstColy(Ly, N))): End Function
Private Function NTermRstColy(Ly$(), N%) As StrColy
Dim Coly()
NTermRstColy.Coly = Coly
End Function


Function WdtyzFstNTerm(NTerm%, L$()) As Integer()
If Si(L) = 0 Then Exit Function
Dim O%(), W%(), I
ReDim O(NTerm - 1)
For Each I In Itr(L)
    W = WdtyzFstNTermL(NTerm, L)
    O = Wdtyz2W(O, W)
Next
WdtyzFstNTerm = O
End Function

Function DyoSplitDot(L$()) As Variant()
Dim I: For Each I In Itr(L)
    PushI DyoSplitDot, SplitDot(I)
Next
End Function

Function NTermRstDy(L$(), N%) As Variant()
Dim I: For Each I In Itr(L)
    PushI NTermRstDy, NTermRst(I, N)
Next
End Function

Function WdtyzFstNTermL(N%, Ln) As Integer()
Dim T: For Each T In FstNTerm(Ln, N)
    PushI WdtyzFstNTermL, Len(T)
Next
End Function

Function Wdtyz2W(W1%(), W2%()) As Integer()
Dim O%(): O = W1
Dim I, J%: For Each I In W2
    If I > O(J) Then O(J) = I
    J = J + 1
Next
Wdtyz2W = O
End Function

Function AliLyzDot(Ly_wi_Dot$()) As String()
AliLyzDot = FmtDy(DyoSplitDot(Ly_wi_Dot))
End Function

Sub BrwDotLy(DotLy$())
Brw AliDotLy(DotLy)
End Sub

Function AliDotLy(DotLy$()) As String()
AliDotLy = FmtDy(DyoDotLy(DotLy), Fmt:=eColSep)
End Function

Function AliDotLyzTwoCol(DotLy$()) As String()
AliDotLyzTwoCol = FmtDy(DyoDotLyzTwoCol(DotLy), Fmt:=eColSep)
End Function

Private Sub AmAli2T__Tst()
Dim L$()
L = Sy("AAA B C D", "L BBB CCC")
Ept = Sy("AAA B   C D", _
         "L   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = AliLyz2T(L)
    C
    Return
End Sub
Private Sub AliLyz3T__Tst()
Dim L$()
L = Sy("AAA B C D", "L BBB CCC")
Ept = Sy("AAA B   C   D", _
         "L   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = AliLyz3T(L)
    C
    Return
End Sub

Function AliTsv(Tsv$()) As String()
AliTsv = FmtDy(AliDy(DyzTsv(Tsv)))
End Function

Function DyzTsv(Tsv$()) As Variant()
Dim L: For Each L In Itr(Tsv)
    PushI DyzTsv, SplitTab(L)
Next
End Function

Function AliLyzSepss(Ly$(), SepSS$) As String()
AliLyzSepss = JnDy(AliDy(LnBrkDy(Ly, SyzSS(SepSS))))
End Function
