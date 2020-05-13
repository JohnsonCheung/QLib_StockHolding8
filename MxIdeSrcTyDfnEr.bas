Attribute VB_Name = "MxIdeSrcTyDfnEr"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcTyDfnEr."
Private Sub IsLnTyDfnEr__Tst()
Dim O$()
Dim L: For Each L In SrczP(CPj)
    If IsLnTyDfnEr(L) Then
        PushI O, L
    End If
Next
Brw O
End Sub
Function IsLnTyDfnEr(Ln) As Boolean
Dim L$: L = Ln
If ShfTyDfnNm(L) = "" Then Exit Function
ShfColonTy L  ' It is optional
L = Trim(L)   ' Then ! ... is must
Select Case True
Case FstChr(L) = "!": Exit Function     '<-- It Valid line
End Select
IsLnTyDfnEr = True     '<-- It is ErLn
End Function

Function TyDfnErLnAyP() As String()
Dim L: For Each L In SrczP(CPj)
    If IsLnTyDfnEr(L) Then
        PushI TyDfnErLnAyP, L
    End If
Next
End Function

Private Sub TyDfnErLnAyP__Tst()
Brw TyDfnErLnAyP
End Sub
