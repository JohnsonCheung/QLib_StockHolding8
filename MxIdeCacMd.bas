Attribute VB_Name = "MxIdeCacMd"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeMdCache."

Function IsCacM() As Boolean
IsCacM = IsCaczM(CMd)
End Function

Function IsCaczMdn(Mdn) As Boolean
IsCaczMdn = IsCaczM(Md(Mdn))
End Function

Function IsCaczM(M As CodeModule) As Boolean ' Is given @M src is eq CacSrc of @M.
Const CSub$ = CMod & "IsCaczM"
Const Trc As Boolean = False
Dim Cs$(): Cs = CacSrczM(M)       ' Cs cached src
Dim Clc&: Clc = Si(Cs)            ' Clc cached line count
Dim Mlc&: Mlc = M.CountOfLines + 1 ' Mlc module line count
If Mlc <> Clc Then
    If Trc Then InfLn CSub, "LnCnt dif", "Md-LnCnt Src-LnCnt", Mlc, Clc
    Exit Function
End If
If Clc = 0 Then IsCaczM = True
Dim Ms$(): Ms = Src(M)         ' Ms module src
Push Ms, ""
IsCaczM = IsEqSy(Cs, Ms)
End Function

Function CacSrcM() As String()
CacSrcM = CacSrczM(CMd)
End Function

Function CacSrczM(M As CodeModule) As String()
Dim F$:     F = SrcFfn(M.Parent)
                If NoFfn(F) Then Exit Function
Dim S$():   S = LyzFt(F)
Dim S1$(): S1 = RmvCls4Sigln(S)
     CacSrczM = RmvAtrVbLn(S1)
End Function

Function RmvAtrVbLn(Src$()) As String()
Dim N&: N = AtrVbLnCnt(Src$())
RmvAtrVbLn = AeFstNEle(Src, N)
End Function

Function AtrVbLnCnt%(Src$())
Dim O%:
    Dim L: For Each L In Itr(Src)
        If NoPfx(L, "Attribute VB") Then Exit For
        O = O + 1
    Next
AtrVbLnCnt = O
End Function

Function RmvCls4Sigln(Src$()) As String()
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
If HasCls4Sigln(Src) Then
    RmvCls4Sigln = AeFstNEle(Src, 4)
Else
    RmvCls4Sigln = Src
End If
End Function

Function HasCls4Sigln(Src$()) As Boolean
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
If Si(Src) < 4 Then Exit Function
If Src(0) <> "VERSION 1.0 CLASS" Then Exit Function
If Src(1) <> "BEGIN" Then Exit Function
If HasPfx(Src(2), "  MultiUse =") Then Exit Function
If Src(3) = "End" Then Exit Function
HasCls4Sigln = True
End Function
