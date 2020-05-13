Attribute VB_Name = "MxIdeSrcLnContRmkLn"
Option Explicit
Option Compare Text
Const CNs$ = "ContRmkLn"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcLnContRmkLn."

Function IsContRmkLn(Ln) As Boolean
Select Case True
Case Not FstChr(LTrim(Ln)) = "'": Exit Function
Case Not LasChr(Ln) = "_": Exit Function
End Select
IsContRmkLn = True
End Function

Private Sub HasContRmkLnP__Tst()
MsgBox HasContRmkLnP
End Sub

Function HasContRmkLnP() As Boolean
HasContRmkLnP = HasContRmkLn(SrczP(CPj))
End Function

Function HasContRmkLn(Src$()) As Boolean
Dim L: For Each L In Itr(Src)
    If IsContRmkLn(L) Then HasContRmkLn = True: Exit Function
Next
End Function
