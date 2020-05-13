Attribute VB_Name = "MxIdeSrcTyDfn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxIdeSrcTyDfn."
Public Const TyDfnFF$ = "Mdn Nm Ty Mem Rmk"

Private Sub TyDfnDrsP__Tst()
BrwDrs TyDfnDrsP
End Sub

Function IsLnTyDfn(Ln) As Boolean
Dim L$: L = Ln
Dim A$: A = ShfTyDfnNm(L): If A = "" Then Exit Function
ShfColonTy L
ShfMemNm L
If L = "" Then IsLnTyDfn = True: Exit Function
If FstChr(L) = "!" Then IsLnTyDfn = True
End Function

Function TyDfnNyP() As String()
TyDfnNyP = TyDfnNy(SrclP)
End Function

Function TyDfnNy(Srcl) As String()
TyDfnNy = TyDfnNyzS(SplitCrLf(Srcl))
End Function

Function TyDfnNyzS(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLnTyDfn(L) Then
        PushI TyDfnNyzS, RmvFstChr(T1(L))
    End If
Next
End Function
Function TyDfnNm$(Ln)
Dim T$: T = T1(Ln)
If T = "" Then Exit Function
If Fst2Chr(T) <> "':" Then Exit Function
If LasChr(T) <> ":" Then Exit Function
TyDfnNm = RmvFstChr(T)
End Function

Function IsLnTyDfnRmk(Ln) As Boolean
If FstChr(Ln) <> "'" Then Exit Function
If FstChr(LTrim(RmvFstChr(Ln))) <> "!" Then Exit Function
IsLnTyDfnRmk = True
End Function

Function IsTyDfnNm(Nm$) As Boolean
':TyDfnNm: :Nm ! #TyDfn-Name# It must be from a str with fst2chr is [':], and then non-space-chr, and then [:].
'              ! Then non-space char is :TyDfnNm
Select Case True
Case Fst2Chr(Nm) <> "':", LasChr(Nm) <> ":"
Case Else: IsTyDfnNm = True
End Select
End Function

Function IsDfnTy(Term$) As Boolean
IsDfnTy = FstChr(Term) = ":"
End Function

Function IsMemNm(Term$) As Boolean
If Len(Term) > 3 Then
    If FstChr(Term) = "#" Then
        If LasChr(Term) = "#" Then
            IsMemNm = True
        End If
    End If
End If
End Function
