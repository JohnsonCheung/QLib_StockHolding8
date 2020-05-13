Attribute VB_Name = "MxIdeSrcLn"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIdeSrcLn."
Sub NRmkBlnkLnAbove__Tst()
Dim M As CodeModule
Dim C As VBComponent: For Each C In CPj.VBComponents
    Set M = C.CodeModule
    Dim N%: N = NRmkBlnkLnAbove(M, M.CountOfDeclarationLines)
    If N > 0 Then Debug.Print Mdn(M), N
Next
End Sub

Function NRmkBlnkLnAbove%(M As CodeModule, Lno&) ' N remark or blank line starting count at @Lno and above
Dim O%
Dim L&: For L = Lno To 1 Step -1
    If Not IsRmkOrBlnkln(M.Lines(L, 1)) Then
        NRmkBlnkLnAbove = O
        Exit Function
    End If
    O = O + 1
Next
End Function
