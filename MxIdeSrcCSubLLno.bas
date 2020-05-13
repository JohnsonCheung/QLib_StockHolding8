Attribute VB_Name = "MxIdeSrcCSubLLno"
Option Explicit
Option Compare Text
Const CNs$ = "Src.CSub"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcCSubLLn."
Function CCSubLLn(L&, Mthly$()) As LLn
Dim Ix&:         Ix = CnstIx(Mthly, "CSub")
                      If Ix < 0 Then Exit Function
        CCSubLLn.Ln = Mthly(Ix)
        CCSubLLn.Lno = L + Ix
End Function

Function EptCSubLLn(L&, Mthly$()) As LLn
If Not IsUsingCSub(Mthly) Then Exit Function
With EptCSubLLn
    .Ln = EptCSubLin(Mthn(Mthly(0)))
    .Lno = NxtSrcIx(Mthly) + L + 1
End With
End Function

Private Function IsUsingCSub(Mthly$()) As Boolean
Const CSub$ = CMod & "IsUsingCSub"
Dim L
IsUsingCSub = True
For Each L In Itr(Mthly)
    If HasSubStr(L, "CSub, ") Then Exit Function
    If HasSubStr(L, "(CSub") Then Exit Function
Next
IsUsingCSub = False
End Function


Private Function CSubLin$(Mthly$(), Mthn$)
If Not IsUsingCSub(Mthly) Then Exit Function
CSubLin = EptCSubLin(Mthn)
End Function

Function EptCSubLin$(Mthn$)
EptCSubLin = FmtQQ("Const CSub$ = CMod & ""?""", Mthn)
End Function

Private Function CSubLno&(Mthly$(), Mthlno&)
Dim I&: I = CnstIx(Mthly, "CSub")
If I = 0 Then Exit Function
CSubLno = I + Mthlno
End Function
