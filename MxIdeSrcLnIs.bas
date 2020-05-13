Attribute VB_Name = "MxIdeSrcLnIs"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcLnIs."
Function IsPrpln(L) As Boolean
IsPrpln = MthKdzL(L) = "Property"
End Function

Private Sub IsMthln__Tst()
GoTo Z
Dim A$
A = "Function IsMthln(A) As Boolean"
Ept = True
GoSub Tst
Exit Sub
Tst:
    Act = IsMthln(A)
    C
    Return
Z:
Dim L, O$()
For Each L In SrcM
    If IsMthln(CStr(L)) Then
        PushI O, L
    End If
Next
Brw O
End Sub

Function IsOptLnOrImplOrBlnk(L) As Boolean
IsOptLnOrImplOrBlnk = True
If IsOptln(L) Then Exit Function
If IsImpln(L) Then Exit Function
If L = "" Then Exit Function
IsOptLnOrImplOrBlnk = False
End Function

Function IsImpln(L) As Boolean
IsImpln = HasPfx(L, "Implements ")
End Function

Function IsOptln(L) As Boolean
If Not HasPfx(L, "Option ") Then Exit Function
Select Case True
Case _
    HasPfx(L, "Option Explicit"), _
    HasPfx(L, "Option Compare Text"), _
    HasPfx(L, "Option Compare Binary"), _
    HasPfx(L, "Option Compare Database")
    IsOptln = True
End Select
End Function

Function IsPubMthln(L) As Boolean
Dim Ln$: Ln = L
Dim Mdy$: Mdy = ShfMdy(Ln): If Mdy <> "" And Mdy <> "Public" Then Exit Function
IsPubMthln = TakMthKd(Ln) <> ""
End Function

Private Sub IsSngMthln__Tst()
Dim L$
GoSub T1
Exit Sub
T1:
    L = "Private Const ShwAllSql$ = SelSql & WhAll & OrdBy"
    Ept = False
    GoTo Tst
Tst:
    Act = IsSngMthln(L)
    C
    Return
End Sub
Function IsSngMthln(L) As Boolean
Dim K$: K = MthKdzL(L): If K = "" Then Exit Function
IsSngMthln = HasSubStr(L, "End " & K)
End Function

Function IsMthln(L) As Boolean
IsMthln = MthTy(L) <> ""
End Function

Function IsUdtln(A) As Boolean
IsUdtln = HasPfx(RmvMdy(A), "Type ")
End Function

Function IsNonSrcLn(L) As Boolean
IsNonSrcLn = True
If HasPfx(L, "Option ") Then Exit Function
Dim Ln$: Ln = Trim(L)
If Ln = "" Then Exit Function
IsNonSrcLn = False
End Function

Function IsSngTermLn(L) As Boolean
IsSngTermLn = InStr(Trim(L), " ") = 0
End Function

Function IsDDLn(L) As Boolean
IsDDLn = Fst2Chr(LTrim(L)) = "--"
End Function

Function IsRmkln(L) As Boolean
IsRmkln = FstChr(LTrim(L)) = "'"
End Function

Function IsRmkOrBlnkln(L) As Boolean
Dim Ln$: Ln = LTrim(L)
Select Case True
Case Ln = "", FstChr(Ln) = "'": IsRmkOrBlnkln = True
End Select
End Function
