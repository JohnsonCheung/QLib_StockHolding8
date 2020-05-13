Attribute VB_Name = "MxIdeSrcContLn"
Option Compare Text
Option Explicit
Const CNs$ = "Contln"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcContLn."
Type TContln
    Ln As String
    Cnt As Integer
End Type

Function TContln(Ln$, Cnt%) As TContln: With TContln: .Ln = Ln: .Cnt = Cnt: End With: End Function
Sub RmvContln(M As CodeModule, Lno&)
M.DeleteLines Lno, NContln(Src(M), Lno)
End Sub

Function IsEqTContln(A As TContln, B As TContln) As Boolean
If A.Cnt = B.Cnt Then
    If A.Ln = B.Ln Then IsEqTContln = True
End If
End Function

Function Contln$(Src$(), Ix)
If Ix = -1 Then Exit Function
Contln = XLn(Src, Ix, NContln(Src, Ix))
End Function

Function TContlnFmSrc(Src$(), Ix) As TContln
Dim Cnt%: Cnt = NContln(Src, Ix)
With TContlnFmSrc
    .Ln = XLn(Src, Ix, Cnt)
    .Cnt = Cnt
End With
End Function


Function Contly(Src$()) As String()
Dim IsContPrv As Boolean, IsCont As Boolean
Dim O$()
    Dim L: For Each L In Itr(Src)
        IsCont = LasChr(L) = "_"
        If IsCont Then L = RmvLasChr(L)
        If IsContPrv Then
            Dim U&: U = UB(O)
            O(U) = O(U) & LTrim(L)
        Else
            PushI O, L
        End If
        IsContPrv = IsCont
    Next
Contly = O
End Function

'--
Private Sub ContSrc__Tst()
Brw ContSrc(SrczP(CPj))
End Sub

Function ContSrc(Src$()) As String()
If Si(Src) = 0 Then Exit Function
Dim O$()
Dim Fst As Boolean: Fst = True
Dim L: For Each L In Itr(Src)
    If Fst Then
        Fst = False
        PushI O, Src(0)
    Else
        If LasChr(L) = "_" Then
            Dim U&: U = UB(O)
            O(U) = RmvLasChr(O(U)) & LTrim(L)
        Else
            PushI O, L
        End If
    End If
Next
ContSrc = O
End Function

Function ContlnzM$(M As CodeModule, Lno&)
If Lno = 0 Then Exit Function
Dim O$
Dim J&: For J = Lno To M.CountOfLines
    Dim L$: L = M.Lines(J, 1)
    If LasChr(L) <> "_" Then
        If O = "" Then
            ContlnzM = L
        Else
            ContlnzM = O & LTrim(L)
        End If
        Exit Function
    End If
    O = O & RmvLasChr(LTrim(L))
Next
Thw CSub, "Las Ln of @Md has [_] at end", "@Md", Mdn(M)
End Function

Function NxtSrcIx&(Src$(), Optional Ix = 0)
Const CSub$ = CMod & "NxtSrcIx"
':NxtSrcIx: :Ix #Nxt-Src-Ix# ! The src-Ln @Src(Ix) maybe a Contln
Dim O&: For O = Ix + 1 To UB(Src)
    If LasChr(Src(O - 1)) <> " _" Then NxtSrcIx = O: Exit Function
Next
Thw CSub, "LasLin of src is a cont-line", "Ix Src", Src
End Function

Private Sub Contln__Tst()
Dim Src$(), Mthix, Act As TContln, Ept As TContln
Mthix = 0
Dim O$(3)
O(0) = "A _"
O(1) = "  B _"
O(2) = "C"
O(3) = "D"
Src = O
Ept = TContln("ABC", 0)
GoSub Tst
Exit Sub
Tst:
    Act = TContlnFmSrc(Src, Mthix)
    Ass IsEqTContln(Act, Ept)
    Return
End Sub
Function ContlnCntzM%(M As CodeModule, Lno)
Const CSub$ = CMod & "ContlnCntzM"
Dim J&, O%
For J = Lno To M.CountOfLines
    O = O + 1
    If LasChr(M.Lines(J, 1)) <> "_" Then
        ContlnCntzM = O
        Exit Function
    End If
Next
Thw CSub, "LasLin of Md cannot be end of [_]", "LasLin-Of-Md Md", M.Lines(M.CountOfLines, 1), Mdn(M)
End Function

Function JnContln$(Contly)
Dim J%, L$, O$()
PushI O, Contly(0)
For J = 1 To UB(Contly) - 1
    PushI O, Contly(J)
Next
End Function

Function ContEIx&(Src$(), Ix&)
Const CSub$ = CMod & "ContEIx"
Dim O&: For O = Ix To UB(Src)
    If LasChr(Src(O)) <> "_" Then ContEIx = O: Exit Function
Next
Thw CSub, "las Ln of @Src has LasChr = '_'", "Las-Src-Ele", LasEle(Src)
End Function

Function NContln(Src$(), Mthix) As Byte
Const CSub$ = CMod & "NContln"
Dim J&, O&
For J = Mthix To UB(Src)
    O = O + 1
    If LasChr(Src(J)) <> "_" Then NContln = O: Exit Function
Next
Thw CSub, "LasEle of Src has LasChr = _", "Src", Src
End Function

Function NxtMdLno(M As CodeModule, Lno)
Const CSub$ = CMod & "NxtMdLno"
Dim J&
For J = Lno To M.CountOfLines
    If LasChr(M.Lines(Lno, 1)) <> "_" Then
        NxtMdLno = J
        Exit Function
    End If
Next
Thw CSub, "All line From Lno has _ as LasChr", "Lno Md Src", Lno, Mdn(M), AmAddIxPfx(Src(M), 1)
End Function

'== X
Private Function XLn$(Src$(), Ix, Cnt%)
If Cnt = 1 Then
    XLn = Src(Ix)
    Exit Function
End If
Dim O$()
    Dim J&: For J = Ix To Ix + Cnt - 2
        PushI O, RmvLasChr(Src(J))
    Next
    PushI O, Src(Ix + Cnt - 1)
XLn = Jn(O)
End Function
